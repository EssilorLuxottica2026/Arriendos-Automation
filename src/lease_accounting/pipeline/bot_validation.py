import csv
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable
import xml.etree.ElementTree as ET

from .pdf_reader import InvoiceSupportReader
from .money import round_cop


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _number(value) -> Decimal:
    result = Decimal(str(value))
    if not result.is_finite():
        raise ValueError("Importe no finito")
    return result


def _same_csv_value(actual: str, expected) -> bool:
    if actual == _text(expected):
        return True
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        try:
            return _number(actual) == _number(expected)
        except (ValueError, InvalidOperation):
            return False
    if isinstance(expected, bool):
        return actual.lower() == str(expected).lower()
    return False


def validate_preparation(
    output_dir: Path,
    postings: list[dict],
    invoice_key: Callable[[str | None], str | None],
    vendor_key: Callable[[str | None], str | None],
    valid_ceco: Callable[[str | None], bool],
) -> dict[str, list[str]]:
    reader = InvoiceSupportReader()
    support_cache = {}
    errors: dict[str, list[str]] = {}
    owners: dict[Path, tuple] = {}
    source_postings: dict[str, list[dict]] = {}
    invoice_sources: dict[str, set[str]] = {}
    for posting in postings:
        folio = str(posting["invoice_id"])
        part_index = str(posting["posting_index"])
        part_errors = errors.setdefault(folio, [])
        prefix = f"Parte {part_index}: "
        if not invoice_key(folio):
            part_errors.append(prefix + "falta un folio valido.")
        for field, label in [("vendor", "vendor"), ("store", "tienda"), ("ceco", "CeCo")]:
            if not _text(posting.get(field)):
                part_errors.append(prefix + f"falta {label}.")
        if not valid_ceco(posting.get("ceco")):
            part_errors.append(prefix + "CeCo invalido.")
        if not vendor_key(posting.get("profit_center")):
            part_errors.append(prefix + "falta profit center.")
        if not posting.get("invoice_date"):
            part_errors.append(prefix + "falta fecha de factura.")
        if posting.get("currency") != "COP":
            part_errors.append(prefix + "moneda distinta de COP.")
        if not posting["allocated_costs"]:
            part_errors.append(prefix + "no hay allocated costs vinculados.")
        try:
            if _number(posting["invoice_total"]) <= 0:
                part_errors.append(prefix + "el importe debe ser mayor que cero.")
            if _number(posting["invoice_total"]) != _number(posting["amount"]) + _number(posting["vat_total"]):
                part_errors.append(prefix + "el importe no coincide con base mas IVA.")
            if _number(posting["amount"]) < 0 or _number(posting["vat_total"]) < 0:
                part_errors.append(prefix + "base o IVA negativo.")
        except (ValueError, InvalidOperation):
            part_errors.append(prefix + "importe, base o IVA invalido.")

        source_path = posting.get("support_source_path")
        if source_path:
            source_postings.setdefault(source_path, []).append(posting)
            invoice_sources.setdefault(folio, set()).add(source_path)
            source = Path(source_path)
            if source.suffix.lower() != ".xml":
                part_errors.append(prefix + "el soporte debe ser XML; no se admiten PDF.")
            else:
                if source_path not in support_cache:
                    try:
                        support_cache[source_path] = reader.parse_file(source)
                    except (OSError, ET.ParseError, ValueError) as exc:
                        support_cache[source_path] = str(exc)
                parsed = support_cache[source_path]
                if isinstance(parsed, str):
                    part_errors.append(prefix + f"no se pudo leer el XML: {parsed}")
                else:
                    if invoice_key(parsed.invoice_id) != invoice_key(folio):
                        part_errors.append(prefix + "el folio del XML no corresponde a la factura.")
                    if parsed.document_type != posting.get("ubl_document_type"):
                        part_errors.append(prefix + "el tipo de documento XML no coincide.")
                    if parsed.invoice_date is None:
                        part_errors.append(prefix + "el XML no tiene fecha de factura.")
                    elif _text(posting.get("invoice_date"))[:10] != parsed.invoice_date.date().isoformat():
                        part_errors.append(prefix + "la fecha de factura no coincide con el XML.")
                    if parsed.barcode and parsed.barcode != posting.get("barcode"):
                        part_errors.append(prefix + "el barcode no coincide con ParentDocumentID del XML.")

        for kind, relative in posting["archivos_csv"].items():
            path = (output_dir / relative).resolve()
            if not path.is_relative_to(output_dir.resolve()):
                raise ValueError(f"El CSV sale del directorio de la ejecucion: {relative}")
            identity = (folio, part_index)
            if path in owners and owners[path] != identity:
                part_errors.append(prefix + "el CSV esta asignado a mas de una parte.")
            owners[path] = identity
            if not path.is_file():
                raise ValueError(f"No existe el CSV vinculado: {relative}")
            if kind == "factura":
                with path.open(encoding="utf-8-sig", newline="") as file:
                    manual_rows = list(csv.reader(file))
                if len(manual_rows) < 2:
                    part_errors.append(prefix + "el CSV de factura esta vacio.")
                else:
                    manual_header = dict(zip(manual_rows[0], manual_rows[1]))
                    if manual_header.get("Invoice Number") != folio:
                        part_errors.append(prefix + "el CSV de factura contiene otro folio.")
                    if not _same_csv_value(manual_header.get("Invoice Total Amount", ""), posting["invoice_total"]):
                        part_errors.append(prefix + "el CSV de factura contiene otro importe.")
                    try:
                        start = manual_rows.index(["CostiRipartiti"]) + 2
                        end = manual_rows.index(["Wht"])
                        if manual_rows[start - 1] != ["S/H", "Amount", "GL Account", "Text", "Vat Code", "Profit Center"]:
                            part_errors.append(prefix + "la estructura de allocated costs del CSV de factura es invalida.")
                        cost_rows = [row for row in manual_rows[start:end] if row]
                        columns = [
                            "D/A$S/H", "Imp$Amount", "Conto$GL Account",
                            "Text$Text", "Cod_Iva$Vat Code", "ProfitCenter$Profit Center",
                        ]
                        if len(cost_rows) != len(posting["allocated_costs"]) or any(
                            len(actual) != len(columns) or any(
                                not _same_csv_value(value, expected.get(column))
                                for value, column in zip(actual, columns)
                            )
                            for actual, expected in zip(cost_rows, posting["allocated_costs"])
                        ):
                            part_errors.append(prefix + "allocated costs del CSV de factura no coincide.")
                    except ValueError:
                        part_errors.append(prefix + "el CSV de factura no tiene sus secciones esperadas.")
                continue
            with path.open(encoding="utf-8-sig", newline="") as file:
                rows = list(csv.DictReader(file))
            expected_rows = posting["allocated_costs"] if kind == "allocated_costs" else posting["header"]
            if len(rows) != len(expected_rows) or any(
                set(actual) != set(expected) or any(
                    not _same_csv_value(actual[column], value) for column, value in expected.items()
                )
                for actual, expected in zip(rows, expected_rows)
            ):
                part_errors.append(prefix + f"el CSV {kind} no coincide con los datos vinculados.")
            if kind == "allocated_costs":
                for row in rows:
                    if invoice_key(row.get("Attribuzione$Assignment")) != invoice_key(folio):
                        part_errors.append(prefix + "allocated costs contiene otro folio.")
                    if vendor_key(row.get("ProfitCenter$Profit Center")) != vendor_key(posting.get("profit_center")):
                        part_errors.append(prefix + "allocated costs tiene otro profit center.")
                    if not _text(row.get("Conto$GL Account")):
                        part_errors.append(prefix + "allocated costs no tiene cuenta contable.")
                    try:
                        _number(row.get("Imp$Amount"))
                    except (ValueError, InvalidOperation):
                        part_errors.append(prefix + "allocated costs contiene un importe invalido.")
            else:
                if len(rows) != 1 or rows[0].get("Invoice Number") != folio:
                    part_errors.append(prefix + "la cabecera no identifica una unica factura.")
                elif (
                    not _same_csv_value(rows[0].get("Invoice Total Amount", ""), posting["invoice_total"])
                    or rows[0].get("SAP_KOSTL") != _text(posting.get("ceco"))
                    or rows[0].get("Currency") != posting.get("currency")
                ):
                    part_errors.append(prefix + "importe, CeCo o moneda de cabecera no corresponde.")

    for folio, sources in invoice_sources.items():
        if len(sources) > 1:
            errors[folio].append("Hay varios XML vinculados al mismo folio; la relacion requiere revision.")

    for source_path, source_rows in source_postings.items():
        parsed = support_cache.get(source_path)
        if parsed is None or isinstance(parsed, str):
            continue
        xml_amounts = [item["amount"] for item in parsed.line_items if item.get("amount") is not None]
        gross = round_cop(sum(xml_amounts)) if xml_amounts else parsed.subtotal
        for column, expected, label in [
            ("gross_amount", gross, "valor bruto"),
            ("vat_total", parsed.detected_iva, "IVA"),
        ]:
            try:
                if expected is None or sum(_number(row.get(column)) for row in source_rows) != _number(expected):
                    for row in source_rows:
                        errors[row["invoice_id"]].append(
                            f"La suma de las partes no coincide con el {label} del XML."
                        )
            except (ValueError, InvalidOperation):
                for row in source_rows:
                    errors[row["invoice_id"]].append(f"No se pudo comprobar el {label} contra el XML.")

    return errors
