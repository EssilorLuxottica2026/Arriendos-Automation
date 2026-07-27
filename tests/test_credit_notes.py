from pathlib import Path
import unittest

import pandas as pd

from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader
from lease_accounting.pipeline.processor import LeaseAccountingPipeline


FIXTURES = Path(__file__).parent / "fixtures"


class CreditNoteTests(unittest.TestCase):
    def test_parses_direct_credit_note_lines(self):
        parsed = InvoiceSupportReader().parse_file(FIXTURES / "credit_note.xml")

        self.assertEqual(parsed.document_type, "CreditNote")
        self.assertEqual(parsed.invoice_id, "FENC17661")
        self.assertEqual(parsed.supplier_id, "900534711")
        self.assertEqual(parsed.subtotal, 1000)
        self.assertEqual(parsed.detected_iva, 190)
        self.assertEqual(parsed.total, 1190)
        self.assertEqual(parsed.item_count_hint, 1)
        self.assertEqual(parsed.line_items[0]["description"], "CUOTA DE ADMINISTRACION")
        self.assertEqual(parsed.line_items[0]["quantity"], 1)

    def test_parses_credit_note_embedded_in_attached_document(self):
        parsed = InvoiceSupportReader().parse_file(FIXTURES / "attached_credit_note.xml")

        self.assertEqual(parsed.document_type, "CreditNote")
        self.assertEqual(parsed.invoice_id, "FENC17661")
        self.assertEqual(parsed.subtotal, 500)
        self.assertEqual(parsed.detected_iva, 0)
        self.assertEqual(parsed.line_items[0]["description"], "FONDO DE IMPREVISTOS")

    def test_normalizes_credit_note_prefix_from_invoices_file(self):
        pipeline = LeaseAccountingPipeline("scratch")

        self.assertEqual(
            pipeline._normalize_invoice_id("NOTA CREDITO FENC17661"),
            "FENC17661",
        )
        self.assertEqual(
            pipeline._invoice_key("NOTA CRÉDITO FENC17661"),
            "FENC17661",
        )

    def test_inverts_all_credit_note_posting_keys(self):
        pipeline = LeaseAccountingPipeline("scratch")
        output_df = pd.DataFrame(
            [
                {
                    "invoice_id": "FENC17661",
                    "vendor": "205766",
                    "store": "T091",
                    "ceco": "6763",
                    "profit_center": "6763",
                    "account": "1245150017",
                    "concept": "GC",
                    "rent_type": "RF",
                    "amount": 1000,
                    "gross_amount": 1000,
                    "discount_total": 100,
                    "posting_discount_total": 100,
                    "payable_rounding": 0,
                    "invoice_total": 1190,
                    "vat_total": 190,
                    "vat_vw": 10,
                    "vat_vq": 10,
                    "vw_percent": 5,
                    "vq_percent": 95,
                    "withholding_tax_amount": 0,
                    "text": "07-2026 GC T091 6763",
                    "invoice_date": pd.Timestamp("2026-07-15"),
                    "due_date": pd.NaT,
                    "payment_terms_text": None,
                    "item_count_hint": 1,
                    "ubl_document_type": "CreditNote",
                    "invoice_line_items": [
                        {
                            "description": "CUOTA DE ADMINISTRACION",
                            "amount": 1000,
                        }
                    ],
                    "invoice_discounts": [],
                    "posting_index": 1,
                    "posting_count": 1,
                    "support_split_index": 1,
                    "support_split_factor": 1,
                    "vat_status": "resolved_value",
                }
            ]
        )

        _, _, lines = pipeline._build_lucy_views(output_df)

        keys_by_type = dict(zip(lines["line_type"], lines["posting_key"]))
        self.assertEqual(keys_by_type["invoice_item"], "50")
        self.assertEqual(keys_by_type["discount"], "40")
        self.assertEqual(keys_by_type["vat_vw"], "50")
        self.assertEqual(keys_by_type["vat_vq"], "40")


if __name__ == "__main__":
    unittest.main()
