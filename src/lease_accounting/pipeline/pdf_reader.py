import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from pypdf import PdfReader

from .money import round_cop


class NonInvoiceUBLDocument(ValueError):
    """Raised when an XML is a UBL response/event rather than an invoice."""

    def __init__(self, filename: str, document_type: str):
        self.filename = filename
        self.document_type = document_type
        super().__init__(f"{filename} is a {document_type} document, not a UBL Invoice")


@dataclass
class ParsedInvoiceSupport:
    source_file: str
    source_type: str
    document_type: str
    invoice_id: str | None
    supplier_id: str | None
    supplier_name: str | None
    invoice_date: pd.Timestamp | None
    due_date: pd.Timestamp | None
    subtotal: float | None
    total: float | None
    detected_iva: float | None
    withholding_tax: float | None
    payment_terms_text: str | None
    item_count_hint: int | None
    line_items: list[dict]
    discounts: list[dict]
    payable_rounding: float | None
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
            "document_type",
            "invoice_id",
            "supplier_id",
            "supplier_name",
            "invoice_date",
            "due_date",
            "subtotal",
            "total",
            "detected_iva",
            "withholding_tax",
            "payment_terms_text",
            "item_count_hint",
            "line_items",
            "discounts",
            "payable_rounding",
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
        subtotal = round_cop(self._extract_money_from_text(text, ["Subtotal"]))
        total = self._extract_money_from_text(
            text,
            ["Total a Pagar", "TOTAL FACTURA", "TOTAL DE LA OPERACION", "TOTAL DE LA OPERACIÓN"],
        )
        total = round_cop(total)
        has_zero_iva_hint = bool(re.search(r"IVA\s*0%|0%\s*[0-9.,]+", text, flags=re.IGNORECASE))
        detected_iva = self._extract_detected_iva_from_text(
            text,
            subtotal=subtotal,
            total=total,
            has_zero_iva_hint=has_zero_iva_hint,
        )
        withholding_tax = round_cop(self._extract_withholding_tax_from_text(text))
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
            document_type="Invoice",
            invoice_id=invoice_id,
            supplier_id=None,
            supplier_name=None,
            invoice_date=invoice_date,
            due_date=due_date,
            subtotal=subtotal,
            total=total,
            detected_iva=detected_iva,
            withholding_tax=withholding_tax,
            payment_terms_text=payment_terms_text,
            item_count_hint=item_count_hint,
            line_items=[],
            discounts=[],
            payable_rounding=None,
            has_zero_iva_hint=has_zero_iva_hint,
            text_excerpt=text[:500],
            flags="|".join(flags) if flags else None,
        )

    def _parse_xml_file(self, path: Path) -> ParsedInvoiceSupport:
        document_root = self._extract_invoice_root_from_xml(path)
        document_type = self._local_name(document_root)
        invoice_id = self._find_text(document_root, "./cbc:ID")
        supplier_id = self._find_text(document_root, ".//cac:AccountingSupplierParty//cbc:CompanyID")
        supplier_name = (
            self._find_text(document_root, ".//cac:AccountingSupplierParty//cac:PartyName/cbc:Name")
            or self._find_text(document_root, ".//cac:AccountingSupplierParty//cbc:RegistrationName")
        )
        invoice_date = self._parse_date_value(self._find_text(document_root, "./cbc:IssueDate"))
        due_date = self._parse_date_value(
            self._find_text(document_root, ".//cac:PaymentMeans/cbc:PaymentDueDate")
            or self._find_text(document_root, "./cbc:DueDate")
        )
        subtotal = round_cop(self._parse_number(self._find_text(document_root, ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount")))
        total = round_cop(self._parse_number(self._find_text(document_root, ".//cac:LegalMonetaryTotal/cbc:PayableAmount")))
        detected_iva = self._resolve_xml_tax_amount(document_root, subtotal, total)
        has_zero_iva_hint = detected_iva == 0
        withholding_tax = self._extract_xml_withholding_tax(document_root)
        payment_terms_text = self._extract_xml_payment_terms(document_root)
        item_count_hint = self._extract_xml_item_count(document_root)
        line_items = self._extract_xml_line_items(document_root)
        line_items.extend(self._extract_xml_document_charges(document_root))
        discounts, unresolved_discount_notes = self._extract_xml_discounts(
            document_root,
            line_items,
            subtotal,
            invoice_date,
        )
        payable_rounding = self._parse_number(
            self._find_text(document_root, ".//cac:LegalMonetaryTotal/cbc:PayableRoundingAmount")
        )
        payable_rounding = round_cop(payable_rounding)
        text_excerpt = self._build_xml_excerpt(document_root)

        flags = []
        if invoice_id is None:
            flags.append("missing_invoice_id")
        if invoice_date is None:
            flags.append("missing_invoice_date")
        if subtotal is None and total is None:
            flags.append("missing_amounts")
        if unresolved_discount_notes:
            flags.append("unresolved_discount_note")

        return ParsedInvoiceSupport(
            source_file=path.name,
            source_type="xml",
            document_type=document_type,
            invoice_id=invoice_id,
            supplier_id=supplier_id,
            supplier_name=supplier_name,
            invoice_date=invoice_date,
            due_date=due_date,
            subtotal=subtotal,
            total=total,
            detected_iva=detected_iva,
            withholding_tax=withholding_tax,
            payment_terms_text=payment_terms_text,
            item_count_hint=item_count_hint,
            line_items=line_items,
            discounts=discounts,
            payable_rounding=payable_rounding,
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
        root_type = self._local_name(root)
        if root_type in {"Invoice", "CreditNote"}:
            return root

        for description in root.findall(f".//{{{self.XML_CBC_NS}}}Description"):
            raw_text = description.text or ""
            for embedded_type in ("Invoice", "CreditNote"):
                match = re.search(
                    rf"(<{embedded_type}\b[\s\S]*</{embedded_type}>)",
                    raw_text,
                )
                if match:
                    return ET.fromstring(match.group(1))

        if root_type == "ApplicationResponse":
            raise NonInvoiceUBLDocument(path.name, root_type)
        raise ValueError(
            f"Could not find embedded Invoice or CreditNote UBL inside {path.name}"
        )

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

    def _resolve_xml_tax_amount(
        self,
        root: ET.Element,
        subtotal: float | None,
        payable_total: float | None,
    ) -> float:
        explicit_tax = self._extract_xml_tax_amount(root)
        if explicit_tax is not None:
            return round_cop(explicit_tax, 0)

        line_extension = self._parse_number(
            self._find_text(root, ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount")
        )
        if line_extension is None:
            line_amounts = [
                self._parse_number(self._find_text(line, "./cbc:LineExtensionAmount"))
                for line in self._xml_document_lines(root)
            ]
            valid_line_amounts = [amount for amount in line_amounts if amount is not None]
            if valid_line_amounts:
                line_extension = round_cop(sum(valid_line_amounts), 0)
        allowance_total = self._parse_number(
            self._find_text(root, ".//cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount")
        )
        charge_total = self._parse_number(
            self._find_text(root, ".//cac:LegalMonetaryTotal/cbc:ChargeTotalAmount")
        )
        tax_exclusive = self._parse_number(
            self._find_text(root, ".//cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount")
        )
        tax_inclusive = self._parse_number(
            self._find_text(root, ".//cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount")
        )

        base_amount = line_extension if line_extension is not None else subtotal
        if base_amount is not None:
            base_amount = base_amount - (allowance_total or 0) + (charge_total or 0)

        # Some vendor XMLs incorrectly report TaxExclusiveAmount as zero even
        # when their invoice lines have a positive net amount.
        if tax_exclusive is not None and (tax_exclusive != 0 or not base_amount):
            base_amount = tax_exclusive

        gross_amount = tax_inclusive if tax_inclusive is not None else payable_total
        if base_amount is not None and gross_amount is not None:
            inferred_tax = round_cop(gross_amount - base_amount, 0)
            if inferred_tax >= 0:
                return max(inferred_tax, 0.0)

        # Business rule: when an XML contains no VAT declaration and its
        # monetary totals cannot provide a positive VAT difference, it is V0.
        return 0.0

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
            return round_cop(sum(values), 0)
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
        lines = self._xml_document_lines(root)
        return len(lines) or None

    def _extract_xml_line_items(self, root: ET.Element) -> list[dict]:
        items = []
        for line in self._xml_document_lines(root):
            line_id = self._find_text(line, "./cbc:ID")
            description = self._extract_xml_line_description(line)
            net_amount = self._parse_number(self._find_text(line, "./cbc:LineExtensionAmount"))
            allowance_amount, charge_amount = self._extract_allowance_charge_totals(line)
            amount = net_amount
            if amount is not None:
                amount = round_cop(amount + allowance_amount - charge_amount, 0)
            net_amount = round_cop(net_amount)
            allowance_amount = round_cop(allowance_amount, 0)
            charge_amount = round_cop(charge_amount, 0)
            tax_amount = self._parse_number(self._find_text(line, ".//cac:TaxTotal/cbc:TaxAmount"))
            tax_amount = round_cop(tax_amount)
            quantity = self._parse_number(
                self._find_text(line, "./cbc:InvoicedQuantity")
                or self._find_text(line, "./cbc:CreditedQuantity")
            )
            if description is None and amount is None:
                continue
            items.append(
                {
                    "line_id": line_id,
                    "description": description,
                    "amount": amount,
                    "net_amount": net_amount,
                    "allowance_amount": allowance_amount,
                    "charge_amount": charge_amount,
                    "tax_amount": tax_amount,
                    "quantity": quantity,
                }
            )
        return items

    def _extract_xml_document_charges(self, root: ET.Element) -> list[dict]:
        parents = {child: parent for parent in root.iter() for child in parent}
        charges = []
        for index, adjustment in enumerate(
            root.findall(".//cac:AllowanceCharge", self.XML_NS),
            start=1,
        ):
            parent = parents.get(adjustment)
            if parent is None or self._local_name(parent) not in {"Invoice", "CreditNote"}:
                continue
            indicator = (self._find_text(adjustment, "./cbc:ChargeIndicator") or "").lower()
            amount = self._parse_number(self._find_text(adjustment, "./cbc:Amount"))
            if indicator != "true" or amount is None or amount <= 0:
                continue
            amount = round_cop(amount, 0)
            reason = (
                self._find_text(adjustment, "./cbc:AllowanceChargeReason")
                or self._find_text(adjustment, "./cbc:AllowanceChargeReasonCode")
                or "CARGO ADICIONAL"
            )
            charge_id = self._find_text(adjustment, "./cbc:ID") or str(index)
            charges.append(
                {
                    "line_id": f"DOCUMENT_CHARGE_{charge_id}",
                    "description": reason,
                    "amount": amount,
                    "net_amount": amount,
                    "allowance_amount": 0,
                    "charge_amount": amount,
                    "tax_amount": 0,
                    "quantity": 1,
                    "source": "xml_document_charge",
                }
            )
        return charges

    def _extract_xml_discounts(
        self,
        root: ET.Element,
        line_items: list[dict],
        subtotal: float | None,
        invoice_date: pd.Timestamp | None,
    ) -> tuple[list[dict], list[str]]:
        parents = {child: parent for parent in root.iter() for child in parent}
        discounts: list[dict] = []
        document_allowance_sum = 0.0

        for allowance in root.findall(".//cac:AllowanceCharge", self.XML_NS):
            charge_indicator = (self._find_text(allowance, "./cbc:ChargeIndicator") or "").lower()
            amount = self._parse_number(self._find_text(allowance, "./cbc:Amount"))
            if charge_indicator != "false" or amount is None or amount <= 0:
                continue

            parent = parents.get(allowance)
            scope = self._local_name(parent) if parent is not None else "Invoice"
            is_line_scope = scope in {"InvoiceLine", "CreditNoteLine"}
            line_id = self._find_text(parent, "./cbc:ID") if is_line_scope else None
            if not is_line_scope:
                document_allowance_sum += amount
            discounts.append(
                {
                    "source": "xml_line_allowance" if is_line_scope else "xml_allowance",
                    "amount": round_cop(amount, 0),
                    "percentage": self._parse_number(
                        self._find_text(allowance, "./cbc:MultiplierFactorNumeric")
                    ),
                    "base_amount": round_cop(self._parse_number(self._find_text(allowance, "./cbc:BaseAmount"))),
                    "reason": self._find_text(allowance, "./cbc:AllowanceChargeReason"),
                    "line_id": line_id,
                }
            )

        legal_allowance = self._parse_number(
            self._find_text(root, ".//cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount")
        )
        if legal_allowance and legal_allowance - document_allowance_sum > 0.01:
            discounts.append(
                {
                    "source": "xml_allowance_total",
                    "amount": round_cop(legal_allowance - document_allowance_sum, 0),
                    "percentage": None,
                    "base_amount": subtotal,
                    "reason": "AllowanceTotalAmount residual",
                    "line_id": None,
                }
            )

        gross_total = self._line_items_total(line_items)
        if gross_total is None:
            gross_total = subtotal

        unresolved_notes: list[str] = []
        seen_notes: set[str] = set()
        # Item descriptions classify the accounting concept; they must never
        # create a second discount merely because their text mentions one.
        note_elements = root.findall(".//cbc:Note", self.XML_NS)
        combined_note_text = " | ".join(
            re.sub(r"\s+", " ", (element.text or "")).strip()
            for element in note_elements
            if (element.text or "").strip()
        )
        combined_discount = self._discount_from_note(
            combined_note_text,
            None,
            line_items,
            gross_total,
            invoice_date,
        )
        combined_key = self._search_key(combined_note_text)
        has_combined_total_discount = bool(
            combined_discount
            and "TOTAL" in combined_key
            and "PAGAR CON DESCUENTO" in combined_key
            and re.search(r"TOTAL[^|]{0,40}PAGAR HASTA", combined_key)
        )
        if has_combined_total_discount:
            discounts.append(combined_discount)

        for note in note_elements:
            raw_note = re.sub(r"\s+", " ", (note.text or "")).strip()
            note_key = self._search_key(raw_note)
            if not note_key or note_key in seen_notes:
                continue
            seen_notes.add(note_key)
            if not re.search(r"INCENTIV|DESCUENT|DCTO|PRONTO PAGO|BONIFIC|REBAJA", note_key):
                continue
            if re.search(r"PERDIDA (?:DE )?DESCUENTO", note_key):
                continue

            if has_combined_total_discount and "PAGAR CON DESCUENTO" in note_key:
                continue

            parent = parents.get(note)
            while parent is not None and self._local_name(parent) not in {
                "InvoiceLine",
                "CreditNoteLine",
            }:
                parent = parents.get(parent)
            line_id = self._find_text(parent, "./cbc:ID") if parent is not None else None
            parsed = self._discount_from_note(raw_note, line_id, line_items, gross_total, invoice_date)
            if parsed:
                discounts.append(parsed)
            elif self._note_has_unchanged_discounted_total(note_key, gross_total):
                continue
            elif not discounts or re.search(r"\d+(?:[.,]\d+)?\s*%|\$\s*[0-9]", note_key):
                unresolved_notes.append(raw_note)

        return discounts, unresolved_notes

    def _note_has_unchanged_discounted_total(self, note_key: str, gross_total: float | None) -> bool:
        if gross_total is None:
            return False
        match = re.search(
            r"(?:TOTAL[^|]{0,40})?PAGAR CON DESCUENTO[^$]{0,80}\$\s*([0-9][0-9.,]+)",
            note_key,
        )
        discounted_total = self._parse_number(match.group(1)) if match else None
        return discounted_total is not None and abs(discounted_total - gross_total) <= 0.01

    def _discount_from_note(
        self,
        note: str,
        line_id: str | None,
        line_items: list[dict],
        gross_total: float | None,
        invoice_date: pd.Timestamp | None,
    ) -> dict | None:
        key = self._search_key(note)
        if gross_total is None or gross_total <= 0:
            return None

        discounted_pair = re.search(
            r"PAGUE CON DESCUENTO.*?VALOR\s*:?\s*([0-9.,]+).*?PAGUE SIN DESCUENTO.*?VALOR\s*:?\s*([0-9.,]+)",
            key,
        )
        if discounted_pair:
            discounted_total = self._parse_number(discounted_pair.group(1))
            regular_total = self._parse_number(discounted_pair.group(2))
            if discounted_total is not None and regular_total is not None and regular_total > discounted_total:
                return self._note_discount(
                    note,
                    regular_total - discounted_total,
                    None,
                    regular_total,
                    invoice_date,
                )

        total_pair = re.search(
            r"TOTAL[^|]{0,40}PAGAR CON DESCUENTO[^$]{0,80}\$\s*([0-9][0-9.,]+).*?"
            r"TOTAL[^|]{0,40}PAGAR HASTA[^$]{0,80}\$\s*([0-9][0-9.,]+)",
            key,
        )
        if total_pair:
            discounted_total = self._parse_number(total_pair.group(1))
            regular_total = self._parse_number(total_pair.group(2))
            if discounted_total is not None and regular_total is not None and regular_total > discounted_total:
                return self._note_discount(
                    note,
                    regular_total - discounted_total,
                    None,
                    regular_total,
                    invoice_date,
                )

        discounted_total_match = re.search(
            r"TOTAL[^|]{0,40}PAGAR CON DESCUENTO[^$]{0,80}\$\s*([0-9][0-9.,]+)",
            key,
        )
        if discounted_total_match:
            discounted_total = self._parse_number(discounted_total_match.group(1))
            if discounted_total is not None and gross_total > discounted_total:
                return self._note_discount(
                    note,
                    gross_total - discounted_total,
                    None,
                    gross_total,
                    invoice_date,
                )

        # A value following "total/pagar con descuento" is the discounted
        # payable total, not the discount amount itself. If no positive
        # difference was established above, the note does not declare an
        # additional discount that can be posted.
        if re.search(
            r"(?:TOTAL|PAGAR)[^|]{0,80}(?:CON|APLICANDO) (?:DESCUENTO|INCENTIVO)"
            r"[^$]{0,80}\$\s*[0-9]",
            key,
        ):
            return None

        explicit_amount = re.search(r"DESCUENTO[^$]{0,180}\$\s*([0-9][0-9.,]+)", key)
        if explicit_amount:
            amount = self._parse_number(explicit_amount.group(1))
            if amount is not None and amount > 0:
                return self._note_discount(note, amount, None, gross_total, invoice_date)

        percent_match = re.search(r"(?:INCENTIVO|DESCUENTO|DESCUENTE)[^%]{0,100}?(\d+(?:[.,]\d+)?)\s*%", key)
        if not percent_match:
            return None
        percentage = self._parse_number(percent_match.group(1))
        if percentage is None or percentage <= 0 or percentage > 100:
            return None

        eligible_items = line_items
        if line_id:
            eligible_items = [item for item in line_items if str(item.get("line_id")) == str(line_id)]
        elif "MODULO GENERAL" in key:
            eligible_items = [
                item for item in line_items if "MODULO GENERAL" in self._search_key(item.get("description"))
            ]
        elif "CUOTA" in key and "ADMINISTRACION" in key and "NO ENTRA FUMIGACION" in key:
            eligible_items = [
                item
                for item in line_items
                if "ADMINISTR" in self._search_key(item.get("description"))
                and "FUMIGACION" not in self._search_key(item.get("description"))
            ]

        base_amount = self._line_items_total(eligible_items)
        if base_amount is None or base_amount <= 0:
            base_amount = gross_total
        amount = round_cop(base_amount * percentage / 100.0, 0)
        return self._note_discount(note, amount, percentage, base_amount, invoice_date)

    def _note_discount(
        self,
        note: str,
        amount: float,
        percentage,
        base_amount,
        invoice_date: pd.Timestamp | None,
    ) -> dict:
        discount = {
            "source": "xml_note",
            "amount": round_cop(amount, 0),
            "percentage": percentage,
            "base_amount": round_cop(base_amount) if base_amount is not None else None,
            "reason": note,
            "line_id": None,
        }
        discount.update(self._extract_discount_payment_condition(note, invoice_date))
        return discount

    def _extract_discount_payment_condition(
        self,
        note: str,
        invoice_date: pd.Timestamp | None,
    ) -> dict:
        key = self._search_key(note)
        condition = {
            "condition_type": "uncalculable",
            "deadline_date": None,
            "deadline_inclusive": True,
            "relative_days": None,
            "business_days": False,
        }

        relative_match = re.search(
            r"(?:DENTRO DE LOS\s+)?(CINCO|DIEZ|\d+)\s+PRIMEROS\s+DIAS(?:\s+(HABILES))?",
            key,
        )
        if relative_match:
            day_token = relative_match.group(1)
            day_count = {"CINCO": 5, "DIEZ": 10}.get(day_token)
            if day_count is None:
                day_count = int(day_token)
            condition.update(
                {
                    "condition_type": "relative_period",
                    "relative_days": day_count,
                    "business_days": bool(relative_match.group(2)),
                }
            )
            return condition

        prefix = r"(ANTES DEL|HASTA(?: EL(?: DIA)?)?)\s+"
        numeric_date = re.search(prefix + r"(20\d{2})[-/.]([01]?\d)[-/.]([0-3]?\d)", key)
        compact_date = re.search(prefix + r"(20\d{2})([01]\d)([0-3]\d)", key)
        deadline = None
        operator = None
        match = numeric_date or compact_date
        if match:
            operator = match.group(1)
            try:
                deadline = pd.Timestamp(int(match.group(2)), int(match.group(3)), int(match.group(4)))
            except ValueError:
                deadline = None

        if deadline is None:
            month_names = {
                "ENERO": 1,
                "FEBRERO": 2,
                "MARZO": 3,
                "ABRIL": 4,
                "MAYO": 5,
                "JUNIO": 6,
                "JULIO": 7,
                "AGOSTO": 8,
                "SEPTIEMBRE": 9,
                "SETIEMBRE": 9,
                "OCTUBRE": 10,
                "NOVIEMBRE": 11,
                "DICIEMBRE": 12,
            }
            natural_date = re.search(
                prefix
                + r"([0-3]?\d)\s+DE\s+("
                + "|".join(month_names)
                + r")(?:\s+DE\s+([0-9.]{4,5}))?",
                key,
            )
            if natural_date:
                operator = natural_date.group(1)
                year_text = (natural_date.group(4) or "").replace(".", "")
                year = int(year_text) if year_text else None
                if year is None and invoice_date is not None and pd.notna(invoice_date):
                    year = int(pd.Timestamp(invoice_date).year)
                if year:
                    try:
                        deadline = pd.Timestamp(
                            year,
                            month_names[natural_date.group(3)],
                            int(natural_date.group(2)),
                        )
                    except ValueError:
                        deadline = None

        if deadline is not None:
            condition.update(
                {
                    "condition_type": "explicit_deadline",
                    "deadline_date": deadline.strftime("%Y-%m-%d"),
                    "deadline_inclusive": operator != "ANTES DEL",
                }
            )
        return condition

    def _extract_allowance_charge_totals(self, container: ET.Element) -> tuple[float, float]:
        allowance_total = 0.0
        charge_total = 0.0
        for adjustment in container.findall("./cac:AllowanceCharge", self.XML_NS):
            amount = self._parse_number(self._find_text(adjustment, "./cbc:Amount"))
            indicator = (self._find_text(adjustment, "./cbc:ChargeIndicator") or "").lower()
            if amount is None:
                continue
            if indicator == "false":
                allowance_total += amount
            elif indicator == "true":
                charge_total += amount
        return round_cop(allowance_total, 0), round_cop(charge_total, 0)

    def _line_items_total(self, line_items: list[dict]) -> float | None:
        amounts = [item.get("amount") for item in line_items if item.get("amount") is not None]
        return round_cop(sum(float(amount) for amount in amounts), 0) if amounts else None

    def _search_key(self, value) -> str:
        text = str(value or "")
        text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
        return re.sub(r"\s+", " ", text.upper()).strip()

    def _local_name(self, element: ET.Element | None) -> str | None:
        if element is None:
            return None
        return element.tag.rsplit("}", 1)[-1]

    def _xml_document_lines(self, root: ET.Element) -> list[ET.Element]:
        return (
            root.findall(".//cac:InvoiceLine", self.XML_NS)
            + root.findall(".//cac:CreditNoteLine", self.XML_NS)
        )

    def _extract_xml_line_description(self, line: ET.Element) -> str | None:
        description = self._find_text(line, ".//cac:Item/cbc:Description")
        if self._is_useful_xml_description(description):
            return description

        notes = [
            (note.text or "").strip()
            for note in line.findall("./cbc:Note", self.XML_NS)
            if (note.text or "").strip()
        ]
        for note in notes:
            budget_match = re.search(
                r"\bPpto:\s*(.*?)(?:\s+Coefic_|\s+Vlr_Ppto_Anual|\s+Div\s+en:|[|/\n\r]|$)",
                note,
                flags=re.IGNORECASE,
            )
            if budget_match:
                candidate = budget_match.group(1).strip()
                if self._is_useful_xml_description(candidate):
                    return candidate
        for note in notes:
            if self._is_useful_xml_description(note):
                return note

        item_id = (
            self._find_text(line, ".//cac:SellersItemIdentification/cbc:ID")
            or self._find_text(line, ".//cac:StandardItemIdentification/cbc:ID")
        )
        return item_id if self._is_useful_xml_description(item_id) else None

    def _is_useful_xml_description(self, value: str | None) -> bool:
        if value is None:
            return False
        text = re.sub(r"\s+", " ", value).strip()
        return bool(text and text not in {".", "-", "--", "_", "N/A", "NA"})

    def _build_xml_excerpt(self, root: ET.Element) -> str:
        parts = []
        for line in self._xml_document_lines(root):
            description = self._find_text(line, "./cac:Item/cbc:Description")
            if description:
                parts.append(description)
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
            return 0
        iva_candidates = re.findall(r"\bIVA\b[^\d]{0,12}([0-9.,]+)", text, flags=re.IGNORECASE)
        parsed = [self._parse_number(item) for item in iva_candidates]
        parsed = [item for item in parsed if item is not None]
        if parsed:
            return round_cop(max(parsed), 0)
        if subtotal is not None and total is not None and total >= subtotal:
            diff = round_cop(total - subtotal, 0)
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
