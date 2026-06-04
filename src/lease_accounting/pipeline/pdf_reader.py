import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from pypdf import PdfReader


@dataclass
class ParsedInvoiceSupport:
    source_file: str
    source_type: str
    invoice_id: str | None
    invoice_date: pd.Timestamp | None
    due_date: pd.Timestamp | None
    subtotal: float | None
    total: float | None
    detected_iva: float | None
    withholding_tax: float | None
    payment_terms_text: str | None
    item_count_hint: int | None
    has_zero_iva_hint: bool
    text_excerpt: str
    flags: str | None


class InvoiceSupportReader:
    XML_CBC_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
    XML_CAC_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
    XML_INV_NS = "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
    XML_NS = {"cbc": XML_CBC_NS, "cac": XML_CAC_NS, "inv": XML_INV_NS}

    INVOICE_PATTERNS = [
        r"Factura Electr[oó]nica De Venta No\.?\s*([A-Z]{1,6}\s*\d{3,})",
        r"FACTURA ELECTR[ÓO]NICA DE\s+VENTA\s+([A-Z]{1,6}\s*\d+)",
        r"\b((?:SM|FA|FAE|FE)\s*\d{3,})\b",
        r"Cantidad\s+Valor Unitario\s+TotalIVA\s+(\d{3,})",
        r"\b(\d{4,})\b",
    ]
    DATE_PATTERNS = [
        (r"Fecha Factura:\s*(\d{4}-\d{2}-\d{2})", False),
        (r"FECHA FACTURA\s*(\d{2}/\d{2}/\d{4})", True),
        (r"Expedici[oó]n\s*(\d{4}-\d{2}-\d{2})", False),
        (r"\bFecha\s+(\d{2}-\d{2}-\d{4})\b", True),
    ]
    PAYMENT_TERMS_BLOCKLIST = {
        "PERIODO",
        "LOCAL",
        "FACTURA",
        "CLIENTE",
        "POR CONCEPTO DE",
        "BANCO CHEQUE NO VALOR PAGADO",
    }
    PAYMENT_MODE_CODE_MAP = {
        "1": "Instrumento no definido",
        "10": "Efectivo",
        "20": "Cheque",
        "30": "Transferencia Crédito",
        "42": "Consignación bancaria",
        "47": "Transferencia Débito Bancaria",
    }

    def parse_many(self, paths: list[Path]) -> pd.DataFrame:
        records = [self.parse_file(path).__dict__ for path in paths]
        columns = [
            "source_file",
            "source_type",
            "invoice_id",
            "invoice_date",
            "due_date",
            "subtotal",
            "total",
            "detected_iva",
            "withholding_tax",
            "payment_terms_text",
            "item_count_hint",
            "has_zero_iva_hint",
            "text_excerpt",
            "flags",
        ]
        if not records:
            return pd.DataFrame(columns=columns)
        df = pd.DataFrame(records)
        df["invoice_key"] = df["invoice_id"].map(self._invoice_key)
        return df

    def parse_file(self, path: Path) -> ParsedInvoiceSupport:
        suffix = path.suffix.lower()
        if suffix == ".xml":
            return self._parse_xml_file(path)
        if suffix == ".pdf":
            return self._parse_pdf_file(path)
        raise ValueError(f"Unsupported invoice support file type: {path.name}")

    def _parse_pdf_file(self, path: Path) -> ParsedInvoiceSupport:
        text = self._extract_pdf_text(path)
        invoice_id = self._extract_invoice_id(text)
        invoice_date = self._extract_invoice_date_from_text(text)
        due_date = self._extract_due_date_from_text(text)
        subtotal = self._extract_money_from_text(text, ["Subtotal"])
        total = self._extract_money_from_text(
            text,
            ["Total a Pagar", "TOTAL FACTURA", "TOTAL DE LA OPERACION", "TOTAL DE LA OPERACIÓN"],
        )
        has_zero_iva_hint = bool(re.search(r"IVA\s*0%|0%\s*[0-9.,]+", text, flags=re.IGNORECASE))
        detected_iva = self._extract_detected_iva_from_text(
            text,
            subtotal=subtotal,
            total=total,
            has_zero_iva_hint=has_zero_iva_hint,
        )
        withholding_tax = self._extract_withholding_tax_from_text(text)
        payment_terms_text = self._extract_payment_terms_from_text(text)
        item_count_hint = self._extract_item_count_hint_from_text(text)

        flags = []
        if invoice_id is None:
            flags.append("missing_invoice_id")
        if invoice_date is None:
            flags.append("missing_invoice_date")
        if subtotal is None and total is None:
            flags.append("missing_amounts")

        return ParsedInvoiceSupport(
            source_file=path.name,
            source_type="pdf",
            invoice_id=invoice_id,
            invoice_date=invoice_date,
            due_date=due_date,
            subtotal=subtotal,
            total=total,
            detected_iva=detected_iva,
            withholding_tax=withholding_tax,
            payment_terms_text=payment_terms_text,
            item_count_hint=item_count_hint,
            has_zero_iva_hint=has_zero_iva_hint,
            text_excerpt=text[:500],
            flags="|".join(flags) if flags else None,
        )

    def _parse_xml_file(self, path: Path) -> ParsedInvoiceSupport:
        invoice_root = self._extract_invoice_root_from_xml(path)
        invoice_id = self._find_text(invoice_root, "./cbc:ID")
        invoice_date = self._parse_date_value(self._find_text(invoice_root, "./cbc:IssueDate"))
        due_date = self._parse_date_value(
            self._find_text(invoice_root, ".//cac:PaymentMeans/cbc:PaymentDueDate")
            or self._find_text(invoice_root, "./cbc:DueDate")
        )
        subtotal = self._parse_number(self._find_text(invoice_root, ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount"))
        total = self._parse_number(self._find_text(invoice_root, ".//cac:LegalMonetaryTotal/cbc:PayableAmount"))
        detected_iva = self._extract_xml_tax_amount(invoice_root)
        has_zero_iva_hint = detected_iva == 0 if detected_iva is not None else False
        withholding_tax = self._extract_xml_withholding_tax(invoice_root)
        payment_terms_text = self._extract_xml_payment_terms(invoice_root)
        item_count_hint = self._extract_xml_item_count(invoice_root)
        text_excerpt = self._build_xml_excerpt(invoice_root)

        flags = []
        if invoice_id is None:
            flags.append("missing_invoice_id")
        if invoice_date is None:
            flags.append("missing_invoice_date")
        if subtotal is None and total is None:
            flags.append("missing_amounts")

        return ParsedInvoiceSupport(
            source_file=path.name,
            source_type="xml",
            invoice_id=invoice_id,
            invoice_date=invoice_date,
            due_date=due_date,
            subtotal=subtotal,
            total=total,
            detected_iva=detected_iva,
            withholding_tax=withholding_tax,
            payment_terms_text=payment_terms_text,
            item_count_hint=item_count_hint,
            has_zero_iva_hint=has_zero_iva_hint,
            text_excerpt=text_excerpt[:500],
            flags="|".join(flags) if flags else None,
        )

    def _extract_pdf_text(self, path: Path) -> str:
        reader = PdfReader(str(path))
        text_parts = []
        for page in reader.pages:
            text_parts.append(page.extract_text() or "")
        text = "\n".join(text_parts)
        return re.sub(r"[ \t]+", " ", text)

    def _extract_invoice_root_from_xml(self, path: Path) -> ET.Element:
        root = ET.parse(path).getroot()
        if root.tag.endswith("Invoice"):
            return root

        description = root.find(f".//{{{self.XML_CBC_NS}}}Description")
        raw_text = description.text if description is not None and description.text else ""
        match = re.search(r"(<Invoice[\s\S]*</Invoice>)", raw_text)
        if not match:
            raise ValueError(f"Could not find embedded Invoice UBL inside {path.name}")
        return ET.fromstring(match.group(1))

    def _find_text(self, root: ET.Element, xpath: str) -> str | None:
        value = root.findtext(xpath, default=None, namespaces=self.XML_NS)
        if value is None:
            return None
        value = value.strip()
        return value or None

    def _extract_xml_tax_amount(self, root: ET.Element) -> float | None:
        candidate_paths = [
            ".//cac:LegalMonetaryTotal/../cac:TaxTotal/cbc:TaxAmount",
            "./cac:TaxTotal/cbc:TaxAmount",
            ".//cac:TaxTotal/cbc:TaxAmount",
        ]
        for path in candidate_paths:
            value = self._parse_number(self._find_text(root, path))
            if value is not None:
                return value
        return None

    def _extract_xml_withholding_tax(self, root: ET.Element) -> float | None:
        explicit_paths = [
            ".//cac:WithholdingTaxTotal/cbc:TaxAmount",
            ".//cac:WithholdingTaxTotal//cbc:TaxAmount",
        ]
        values = []
        for path in explicit_paths:
            for element in root.findall(path, self.XML_NS):
                parsed = self._parse_number((element.text or "").strip())
                if parsed is not None:
                    values.append(parsed)
        if values:
            return round(sum(values), 2)
        return None

    def _extract_xml_payment_terms(self, root: ET.Element) -> str | None:
        payment_id = self._find_text(root, ".//cac:PaymentMeans/cbc:PaymentID")
        if payment_id and self._looks_like_real_payment_term(payment_id):
            return payment_id

        note_terms = [
            self._find_text(root, "./cbc:DueDate"),
            self._find_text(root, ".//cac:PaymentMeans/cbc:PaymentDueDate"),
        ]
        if any(note_terms):
            return None

        payment_code = self._find_text(root, ".//cac:PaymentMeans/cbc:PaymentMeansCode")
        if payment_code and payment_code in self.PAYMENT_MODE_CODE_MAP:
            mapped = self.PAYMENT_MODE_CODE_MAP[payment_code]
            if mapped != "Instrumento no definido":
                return mapped
        return None

    def _extract_xml_item_count(self, root: ET.Element) -> int | None:
        line_count = self._find_text(root, "./cbc:LineCountNumeric")
        if line_count and line_count.isdigit():
            return int(line_count)
        lines = root.findall(".//cac:InvoiceLine", self.XML_NS)
        return len(lines) or None

    def _build_xml_excerpt(self, root: ET.Element) -> str:
        parts = []
        for item in root.findall(".//cac:InvoiceLine/cac:Item/cbc:Description", self.XML_NS):
            text = (item.text or "").strip()
            if text:
                parts.append(text)
        if not parts:
            for note in root.findall("./cbc:Note", self.XML_NS):
                text = (note.text or "").strip()
                if text:
                    parts.append(text)
        return " | ".join(parts)

    def _extract_invoice_id(self, text: str) -> str | None:
        for pattern in self.INVOICE_PATTERNS:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                candidate = re.sub(r"\s+", "", match.group(1).upper())
                if candidate in {"LEY1231", "MARZO2026", "CLIENTE", "FACTURA"}:
                    continue
                return candidate
        return None

    def _extract_invoice_date_from_text(self, text: str) -> pd.Timestamp | None:
        for pattern, dayfirst in self.DATE_PATTERNS:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if not match:
                continue
            parsed = self._parse_date_value(match.group(1).strip(), dayfirst=dayfirst)
            if parsed is not None:
                return parsed
        return None

    def _extract_due_date_from_text(self, text: str) -> pd.Timestamp | None:
        patterns = [
            (r"Fecha Vencimiento:\s*(\d{4}-\d{2}-\d{2})", False),
            (r"Fecha de Vencimiento:\s*(\d{4}-\d{2}-\d{2})", False),
            (r"Fecha Vencimiento:\s*(\d{2}-\d{2}-\d{4})", True),
            (r"Fecha de Vencimiento:\s*(\d{2}-\d{2}-\d{4})", True),
            (r"Fecha Vencimiento:\s*(\d{2}/\d{2}/\d{4})", True),
            (r"Fecha de Vencimiento:\s*(\d{2}/\d{2}/\d{4})", True),
            (r"Fecha de Vencimiento:\s*(\d{8})", False),
        ]
        for pattern, dayfirst in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if not match:
                continue
            raw = match.group(1).strip()
            if re.fullmatch(r"\d{8}", raw):
                raw = f"{raw[0:4]}-{raw[4:6]}-{raw[6:8]}"
                dayfirst = False
            parsed = self._parse_date_value(raw, dayfirst=dayfirst)
            if parsed is not None:
                return parsed
        return None

    def _extract_money_from_text(self, text: str, labels: list[str]) -> float | None:
        for label in labels:
            pattern = rf"{label}\s*\$?\s*([0-9][0-9.,]*)"
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                value = self._parse_number(match.group(1))
                if value is not None:
                    return value
        return None

    def _extract_detected_iva_from_text(
        self,
        text: str,
        subtotal: float | None,
        total: float | None,
        has_zero_iva_hint: bool,
    ) -> float | None:
        if has_zero_iva_hint:
            return 0.0
        iva_candidates = re.findall(r"\bIVA\b[^\d]{0,12}([0-9.,]+)", text, flags=re.IGNORECASE)
        parsed = [self._parse_number(item) for item in iva_candidates]
        parsed = [item for item in parsed if item is not None]
        if parsed:
            return max(parsed)
        if subtotal is not None and total is not None and total >= subtotal:
            diff = round(total - subtotal, 2)
            if 0 <= diff <= total:
                return diff
        return None

    def _extract_withholding_tax_from_text(self, text: str) -> float | None:
        patterns = [
            r"Total Impuestos Retenidos\s*\$?\s*([0-9.,]+)",
            r"Withholding Tax Amount\s*\$?\s*([0-9.,]+)",
            r"\bRetefuente\b\s*[:$]\s*([0-9.,]+)",
            r"\bRete\s+fuente\b\s*[:$]\s*([0-9.,]+)",
            r"\bRetención en la fuente\b\s*[:$]\s*([0-9.,]+)",
            r"\bReteica\b\s*[:$]\s*([0-9.,]+)",
            r"\bRete ICA\b\s*[:$]\s*([0-9.,]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                return self._parse_number(match.group(1))
        return None

    def _extract_payment_terms_from_text(self, text: str) -> str | None:
        patterns = [
            r"Cr[eé]dito\s+Plazo\s+(\d+\s*d[ií]as?)",
            r"Plazo\s+(\d+\s*d[ií]as?)",
            r"Forma de Pago:\s*([A-Z0-9 ]{3,40})",
            r"Forma de Pago\s+([A-Z0-9 ]{3,40})",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if not match:
                continue
            value = re.sub(r"\s+", " ", match.group(1)).strip()
            if not self._looks_like_real_payment_term(value):
                continue
            return value
        return None

    def _extract_item_count_hint_from_text(self, text: str) -> int | None:
        numbered_items = {token.lstrip("0") or "0" for token in re.findall(r"\b(00\d|0\d\d)\b", text)}
        if len(numbered_items) >= 2:
            return len(numbered_items)

        line_items = re.findall(r"\b\d+\s+zz\b", text, flags=re.IGNORECASE)
        if len(line_items) >= 2:
            return len(line_items)

        return None

    def _looks_like_real_payment_term(self, value: str) -> bool:
        cleaned = re.sub(r"\s+", " ", value).strip()
        if not cleaned:
            return False
        if cleaned.upper() in self.PAYMENT_TERMS_BLOCKLIST:
            return False
        if re.fullmatch(r"\d+", cleaned):
            return False
        return bool(
            re.search(
                r"DIAS|D[IÍ]AS|CREDITO|CR[EÉ]DITO|CONTADO|TRANSFERENCIA|CHEQUE|EFECTIVO|CONSIGNACI[OÓ]N|PSE",
                cleaned,
                flags=re.IGNORECASE,
            )
        )

    def _parse_date_value(self, raw: str | None, dayfirst: bool = False) -> pd.Timestamp | None:
        if not raw:
            return None
        parsed = pd.to_datetime(raw, errors="coerce", dayfirst=dayfirst)
        if pd.notna(parsed):
            return parsed.normalize()
        return None

    def _parse_number(self, raw: str | None) -> float | None:
        if raw is None:
            return None
        text = raw.strip().replace("$", "").replace(" ", "")
        if not text:
            return None
        if text.count(",") > 1 and "." not in text:
            text = text.replace(",", "")
        elif text.count(".") > 1 and "," not in text:
            text = text.replace(".", "")
        elif "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        else:
            text = text.replace(",", "")
        try:
            return float(text)
        except ValueError:
            return None

    def _invoice_key(self, value: str | None) -> str | None:
        if not value:
            return None
        return re.sub(r"[^A-Z0-9]", "", value.upper())
