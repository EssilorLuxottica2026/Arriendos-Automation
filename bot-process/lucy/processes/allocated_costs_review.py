"""Revisión de Allocated Costs en INTERACTIVE EXPORT después de importar el CSV.

Revisa vendor, importes, cuentas, CeCo, concepto/texto de cada línea, IVA e IVA deducible,
y el resumen de Allocated Cost (campo ``CostiRipartiti``). Si el resumen viene vacío lo
completa con ``MM-YYYY SIGLA1/SIGLA2 TIENDA CECO``. No pulsa Save Document.
"""
import csv
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError, Locator, Page, expect

from ..logging import log
from ..notifications import check_lucy_errors

SH = "D/A$S/H"
AMOUNT = "Imp$Amount"
VAT_CODE = "Cod_Iva$Vat Code"
GL_ACCOUNT = "Conto$GL Account"
COST_CENTER = "CdC$Cost Center"
PROFIT_CENTER = "ProfitCenter$Profit Center"
TEXT = "Text$Text"
BRAND = "Brand$Brand"

DEBIT = "40"
CREDIT = "50"
DEDUCTIBLE_VAT_CODE = "VW"
CONCEPT_SEPARATOR = "/"
PERIOD = re.compile(r"^\d{2}-\d{4}$")
SPLIT_PERCENT = re.compile(r"^\d+(?:[.,]\d+)?%$")
GL_ACCOUNT_FORMAT = re.compile(r"^\d{10}$")

INVOICE_TOTAL_LABEL = "Invoice Total Amount"
TAXABLE_LABEL = "Total Taxable Amount"
VAT_LABEL = "Total VAT Amount"
SUMMARY_INPUT = "#q1-text-input-CostiRipartiti"
GRID_ROW = ".SpreadsheetGrid__row"
GRID_CELL = "div[class*='SpreadsheetGridCell']"
GRID_LOAD_TIMEOUT_MS = 30_000


class AllocatedCostsReviewError(RuntimeError):
    pass


def parse_amount(value) -> Decimal:
    """Importe de Lucy o del CSV: ``48979``, ``48.979``, ``48.979,00``, ``$ 48,979.00``."""
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    text = re.sub(r"[^\d.,-]", "", str(value or ""))
    if not re.search(r"\d", text):
        raise ValueError(f"Importe sin valor numérico: '{value}'.")
    if "." in text and "," in text:
        decimal_sep = "." if text.rfind(".") > text.rfind(",") else ","
        thousands_sep = "," if decimal_sep == "." else "."
        text = text.replace(thousands_sep, "").replace(decimal_sep, ".")
    else:
        sep = "." if "." in text else "," if "," in text else None
        if sep:
            whole, _, tail = text.rpartition(sep)
            if text.count(sep) > 1 or len(tail) == 3:
                text = text.replace(sep, "")
            else:
                text = f"{whole.replace(sep, '')}.{tail}"
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Importe no válido: '{value}'.") from exc


def clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none"} else text


def read_csv_rows(csv_path: Path) -> list[dict]:
    with csv_path.open(encoding="utf-8-sig", newline="") as handle:
        rows = [{key: clean(value) for key, value in row.items()} for row in csv.DictReader(handle)]
    rows = [row for row in rows if any(row.values())]
    if not rows:
        raise AllocatedCostsReviewError(f"El CSV {csv_path.name} no tiene líneas de Allocated Costs.")
    return rows


def line_cost_center(row: dict) -> str:
    return clean(row.get(COST_CENTER)) or clean(row.get(PROFIT_CENTER))


def strip_suffix(text: str, suffix: str) -> str:
    if suffix and (text == suffix or text.endswith(f" {suffix}")):
        return text[: -len(suffix)].rstrip()
    return text


def split_period(text: str) -> tuple[str, str]:
    """Separa ``MM-YYYY`` (y el porcentaje de reparto opcional, ``MM-YYYY 33%``) del resto del texto."""
    period, _, rest = text.partition(" ")
    if not PERIOD.match(period):
        return "", text
    percent, _, after_percent = rest.partition(" ")
    if SPLIT_PERCENT.match(percent):
        return f"{period} {percent}", after_percent
    return period, rest


def split_text(row: dict) -> tuple[str, str]:
    """(periodo, concepto) del texto ``MM-YYYY [NN%] CONCEPTO TIENDA CECO``; la tienda puede tener espacios."""
    text = " ".join(clean(row.get(TEXT)).split())
    period, rest = split_period(text)
    for tail in (line_cost_center(row), " ".join(clean(row.get(BRAND)).split())):
        rest = strip_suffix(rest, tail)
    return period, rest


def line_concept(row: dict) -> str:
    return split_text(row)[1]


def unique(values) -> list[str]:
    seen: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.append(value)
    return seen


def single_value(rows: list[dict], getter, name: str) -> str:
    values = unique(getter(row) for row in rows)
    if len(values) != 1:
        raise AllocatedCostsReviewError(
            f"Las líneas deben compartir un solo {name} para el resumen y tienen: {values or 'ninguno'}."
        )
    return values[0]


def concept_summary(rows: list[dict]) -> str:
    """Resumen de Allocated Cost: ``10-2026 RV/INTERES T004 6554`` (siglas sin repetir)."""
    period = single_value(rows, lambda row: split_text(row)[0], "periodo")
    store = single_value(rows, lambda row: clean(row.get(BRAND)), "tienda")
    cost_center = single_value(rows, line_cost_center, "CeCo")
    concepts = unique(line_concept(row) for row in rows)
    if not concepts:
        raise AllocatedCostsReviewError("Las líneas no tienen concepto en el texto.")
    return f"{period} {CONCEPT_SEPARATOR.join(concepts)} {store} {cost_center}"


def summary_concepts(summary: str, rows: list[dict]) -> list[str]:
    text = " ".join(summary.split())
    _, rest = split_period(text)
    for tails in (unique(line_cost_center(row) for row in rows), unique(" ".join(clean(row.get(BRAND)).split()) for row in rows)):
        if len(tails) == 1:
            rest = strip_suffix(rest, tails[0])
    return [concept.strip() for concept in rest.split(CONCEPT_SEPARATOR) if concept.strip()]


def summary_issues(summary: str, rows: list[dict]) -> list[str]:
    found = summary_concepts(summary, rows)
    issues = []
    missing = [concept for concept in unique(line_concept(row) for row in rows) if concept not in found]
    if missing:
        issues.append(f"al resumen le faltan conceptos: {', '.join(missing)}")
    repeated = unique(concept for concept in found if found.count(concept) > 1)
    if repeated:
        issues.append(f"el resumen repite conceptos: {', '.join(repeated)}")
    return issues


def signed_amount(row: dict) -> Decimal:
    amount = parse_amount(row.get(AMOUNT))
    return -amount if clean(row.get(SH)) == CREDIT else amount


def deductible_vat(rows: list[dict]) -> Decimal:
    """IVA deducible: líneas VW al debe (40) que se compensan con una línea 50 en la misma cuenta."""
    credit_accounts = {clean(row.get(GL_ACCOUNT)) for row in rows if clean(row.get(SH)) == CREDIT}
    return sum(
        (
            parse_amount(row.get(AMOUNT))
            for row in rows
            if clean(row.get(SH)) == DEBIT
            and clean(row.get(VAT_CODE)) == DEDUCTIBLE_VAT_CODE
            and clean(row.get(GL_ACCOUNT)) in credit_accounts
        ),
        Decimal(0),
    )


def accounting_issues(rows: list[dict], part: dict) -> list[str]:
    """Revisión de los datos del CSV contra la parte del control, sin tocar Lucy."""
    issues = []
    expected_ceco = clean(part.get("ceco"))
    expected_store = clean(part.get("tienda"))
    for number, row in enumerate(rows, start=1):
        where = f"línea {number}"
        if clean(row.get(SH)) not in {DEBIT, CREDIT}:
            issues.append(f"{where}: S/H '{row.get(SH)}' no es 40 ni 50")
        try:
            if parse_amount(row.get(AMOUNT)) <= 0:
                issues.append(f"{where}: importe no positivo ({row.get(AMOUNT)})")
        except ValueError as exc:
            issues.append(f"{where}: {exc}")
        if not GL_ACCOUNT_FORMAT.match(clean(row.get(GL_ACCOUNT))):
            issues.append(f"{where}: cuenta '{row.get(GL_ACCOUNT)}' no válida")
        if not clean(row.get(VAT_CODE)):
            issues.append(f"{where}: sin Vat Code")
        if expected_ceco and line_cost_center(row) != expected_ceco:
            issues.append(f"{where}: CeCo '{line_cost_center(row)}' distinto al del control ({expected_ceco})")
        if expected_store and clean(row.get(BRAND)) != expected_store:
            issues.append(f"{where}: tienda '{row.get(BRAND)}' distinta a la del control ({expected_store})")
        if not clean(row.get(TEXT)):
            issues.append(f"{where}: sin texto")
        elif not line_concept(row):
            issues.append(f"{where}: el texto '{row.get(TEXT)}' no tiene concepto")
    if issues:
        return issues

    if part.get("base") is not None:
        net = sum((signed_amount(row) for row in rows), Decimal(0))
        base = parse_amount(part["base"])
        if net != base:
            issues.append(f"el neto de las líneas ({net}) no coincide con la base del control ({base})")

    unmatched = [dict(row) for row in rows if clean(row.get(SH)) == CREDIT]
    for row in rows:
        if clean(row.get(SH)) != DEBIT or clean(row.get(VAT_CODE)) != DEDUCTIBLE_VAT_CODE:
            continue
        match = next(
            (
                credit
                for credit in unmatched
                if clean(credit.get(GL_ACCOUNT)) == clean(row.get(GL_ACCOUNT))
                and parse_amount(credit.get(AMOUNT)) == parse_amount(row.get(AMOUNT))
            ),
            None,
        )
        if match is not None:
            unmatched.remove(match)
    for credit in unmatched:
        issues.append(
            f"la línea 50 de {credit.get(AMOUNT)} en {credit.get(GL_ACCOUNT)} no compensa ningún IVA deducible (VW)"
        )

    vat = deductible_vat(rows)
    if vat and part.get("iva") is not None and vat > parse_amount(part["iva"]):
        issues.append(f"el IVA deducible ({vat}) supera el IVA de la factura ({part['iva']})")
    return issues


def cell_matches(cell: str, expected: str) -> bool:
    cell = clean(cell)
    return cell == expected or cell.startswith(f"{expected} ")


def amount_matches(cell: str, expected: Decimal) -> bool:
    try:
        return parse_amount(cell) == expected
    except ValueError:
        return False


def row_matches(cells: list[str], expected: dict) -> bool:
    for key in (SH, VAT_CODE, GL_ACCOUNT, PROFIT_CENTER, COST_CENTER):
        value = clean(expected.get(key))
        if value and not any(cell_matches(cell, value) for cell in cells):
            return False
    text = clean(expected.get(TEXT))
    if text and not any(clean(cell) == text for cell in cells):
        return False
    amount = parse_amount(expected.get(AMOUNT))
    return any(amount_matches(cell, amount) for cell in cells)


def grid_issues(expected_rows: list[dict], grid_rows: list[list[str]]) -> list[str]:
    """Cada línea del CSV debe estar en Lucy con su importe, cuenta, CeCo, Vat Code y texto."""
    issues = []
    if len(grid_rows) != len(expected_rows):
        issues.append(f"Lucy muestra {len(grid_rows)} líneas y el CSV tiene {len(expected_rows)}")
    available = list(range(len(grid_rows)))
    for number, row in enumerate(expected_rows, start=1):
        match = next((index for index in available if row_matches(grid_rows[index], row)), None)
        if match is None:
            issues.append(
                f"línea {number} ({row.get(SH)} {row.get(AMOUNT)} {row.get(GL_ACCOUNT)} "
                f"'{row.get(TEXT)}') no aparece igual en Lucy"
            )
        else:
            available.remove(match)
    return issues


def invoice_total(invoice: dict, key: str):
    parts = invoice.get("partes") or []
    values = [part.get(key) for part in parts]
    if values and all(value is not None for value in values):
        return sum((parse_amount(value) for value in values), Decimal(0))
    if key == "importe" and invoice.get("importe") is not None:
        return parse_amount(invoice["importe"])
    return None


def read_grid_rows(page: Page) -> list[list[str]]:
    rows = page.locator(GRID_ROW)
    expect(rows.first).to_be_visible(timeout=GRID_LOAD_TIMEOUT_MS)
    result = []
    for index in range(rows.count()):
        cells = [clean(text) for text in rows.nth(index).locator(GRID_CELL).all_inner_texts()]
        if any(cells):
            result.append(cells)
    return result


def read_labeled_amount(page: Page, label: str) -> Decimal:
    from .export_document import labeled_field

    field: Locator = labeled_field(page, label).locator("input").first
    expect(field).to_be_visible()
    return parse_amount(field.input_value())


def header_issues(page: Page, invoice: dict) -> list[str]:
    issues = []
    for label, key in ((INVOICE_TOTAL_LABEL, "importe"), (TAXABLE_LABEL, "base"), (VAT_LABEL, "iva")):
        expected = invoice_total(invoice, key)
        if expected is None:
            continue
        try:
            shown = read_labeled_amount(page, label)
        except (PlaywrightError, AssertionError, ValueError, RuntimeError) as exc:
            issues.append(f"no se pudo leer {label}: {exc}")
            continue
        if shown != expected:
            issues.append(f"{label} en Lucy es {shown} y en el control {expected}")
    return issues


def review_summary(page: Page, folio: str, rows: list[dict]) -> dict:
    """Construye siempre el resumen; si Lucy lo trae lo comprueba y si no lo escribe.

    Devuelve lo que se registra en el JSON: ``resumen``, ``conceptos`` y ``origen``
    (``completado`` o ``ya_existia``).
    """
    expected = concept_summary(rows)
    field = page.locator(SUMMARY_INPUT)
    expect(field).to_be_visible()
    current = clean(field.input_value())
    if current:
        log(f"[{folio}] El resumen de Allocated Cost ya viene en Lucy: '{current}'. Comprobando.")
        issues = summary_issues(current, rows)
        if issues:
            raise AllocatedCostsReviewError(f"Resumen de Allocated Cost incorrecto: {'; '.join(issues)}.")
        log(f"[{folio}] Resumen de Allocated Cost correcto.")
        return {"resumen": current, "conceptos": summary_concepts(current, rows), "origen": "ya_existia"}
    log(f"[{folio}] Resumen de Allocated Cost vacío; completando con '{expected}'.")
    field.fill(expected)
    expect(field).to_have_value(expected)
    check_lucy_errors(page, folio, "completar el resumen de Allocated Cost")
    log(f"[{folio}] Resumen de Allocated Cost completado.")
    return {"resumen": expected, "conceptos": summary_concepts(expected, rows), "origen": "completado"}


def check_before_import(folio: str, part: dict, csv_path: Path) -> list[dict]:
    """Revisa el CSV contra el control antes de tocar Lucy; devuelve sus líneas."""
    rows = read_csv_rows(csv_path)
    issues = accounting_issues(rows, part)
    if issues:
        raise AllocatedCostsReviewError(f"El CSV {csv_path.name} no cuadra con el control: {'; '.join(issues)}.")
    vat = deductible_vat(rows)
    log(
        f"[{folio}] CSV revisado: {len(rows)} líneas, cuentas y CeCo correctos, neto = base"
        + (f", IVA deducible {vat}." if vat else ", sin IVA deducible.")
    )
    concept_summary(rows)
    return rows


def review_allocated_costs(page: Page, folio: str, invoice: dict, rows: list[dict], zero_balance) -> dict:
    """Revisión en Lucy después de Balance:0. ``zero_balance(page, folio)`` vuelve a exigir Balance:0."""
    log(f"[{folio}] Revisando las líneas de Allocated Costs en Lucy.")
    issues = grid_issues(rows, read_grid_rows(page))
    if issues:
        raise AllocatedCostsReviewError(f"Allocated Costs en Lucy no coincide con el CSV: {'; '.join(issues)}.")
    log(f"[{folio}] Las {len(rows)} líneas conservan importe, cuenta, CeCo, Vat Code, concepto y texto.")

    issues = header_issues(page, invoice)
    if issues:
        raise AllocatedCostsReviewError(f"Importes de la factura no cuadran: {'; '.join(issues)}.")
    log(f"[{folio}] Invoice Total, Total Taxable y Total VAT coinciden con el control.")

    summary = review_summary(page, folio, rows)
    zero_balance(page, folio)
    check_lucy_errors(page, folio, "revisar Allocated Costs")
    log(f"[{folio}] La factura cuadra (Balance:0) y Lucy no muestra diferencias ni errores.")
    return summary
