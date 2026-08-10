import unittest
from unittest.mock import patch

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class GcContractAccountTests(unittest.TestCase):
    def setUp(self):
        self.pipeline = LeaseAccountingPipeline("scratch")
        self.item = {"description": "CUOTA DE ADMINISTRACION", "amount": 1000}
        self.pipeline.learning_dictionary = pd.DataFrame(
            [
                {
                    "concept": "GC",
                    "account": "1537210002",
                    "phrases": "CUOTA DE ADMINISTRACION",
                    "text": "GC",
                }
            ]
        )

    def test_positive_sell_media_uses_fixed_gc_account(self):
        concept, account = self.pipeline._item_line_classification(
            self.item,
            {"account": "1537210002", "rent_type": "RV", "sell_media": 125000},
        )

        self.assertEqual(concept, "GC")
        self.assertEqual(account, "1245150017")

    def test_user_learned_gc_selection_still_uses_contract_account(self):
        self.pipeline.learning_dictionary = pd.DataFrame(
            [
                {
                    "concept": "GC",
                    "account": "1537210002",
                    "phrases": "ADM GC",
                    "text": "GC",
                }
            ]
        )
        concept, account = self.pipeline._item_line_classification(
            {"description": "ADM GC % 01.06-30.06.2026", "amount": 165196},
            {"account": "1537210002", "rent_type": "RV", "sell_media": 1545851},
        )

        self.assertEqual(concept, "GC")
        self.assertEqual(account, "1245150017")

    def test_blank_dash_zero_or_missing_sell_media_use_variable_gc(self):
        for sell_media in [None, "-", "", 0]:
            with self.subTest(sell_media=sell_media):
                concept, account = self.pipeline._item_line_classification(
                    self.item,
                    {"account": "1245150017", "rent_type": "RF", "sell_media": sell_media},
                )
                self.assertEqual(concept, "GC VARIABLE")
                self.assertEqual(account, "1537210002")

    def test_allocated_cost_text_uses_gc_variable(self):
        output_df = pd.DataFrame(
            [
                {
                    "invoice_id": "GC1",
                    "vendor": "100",
                    "store": "T001",
                    "ceco": "6001",
                    "profit_center": "6001",
                    "account": "1245150017",
                    "concept": "RF",
                    "rent_type": "RF",
                    "sell_media": 0,
                    "amount": 1000,
                    "gross_amount": 1000,
                    "discount_total": 0,
                    "posting_discount_total": 0,
                    "payable_rounding": 0,
                    "invoice_total": 1000,
                    "vat_total": 0,
                    "vat_vw": 0,
                    "vat_vq": 0,
                    "vw_percent": 1,
                    "vq_percent": 0,
                    "withholding_tax_amount": 0,
                    "text": "08-2026 RF T001 6001",
                    "invoice_date": pd.Timestamp("2026-08-01"),
                    "due_date": pd.NaT,
                    "payment_terms_text": None,
                    "item_count_hint": 1,
                    "ubl_document_type": "Invoice",
                    "invoice_line_items": [self.item],
                    "invoice_discounts": [],
                    "support_split_index": 1,
                    "support_split_factor": 1,
                }
            ]
        )

        _, _, lines = self.pipeline._build_lucy_views(output_df)
        item_line = lines.loc[lines["line_type"].eq("invoice_item")].iloc[0]

        self.assertEqual(item_line["concept"], "GC VARIABLE")
        self.assertEqual(item_line["account"], "1537210002")
        self.assertIn("GC VARIABLE", item_line["text"])

    def test_contract_loader_preserves_sell_media_column(self):
        preview = pd.DataFrame(
            [
                [None, None, None, None, None],
                ["CeCo", "Contract", "End of Term", "Status en REM", "SELL MEDIA"],
            ]
        )
        contract_rows = pd.DataFrame(
            [
                {
                    "CeCo": "6001",
                    "Contract": "C-1",
                    "End of Term": "2026-12-31",
                    "Status en REM": "VIGENTE",
                    "SELL MEDIA": 250000,
                }
            ]
        )

        class FakeExcelFile:
            sheet_names = ["COLOMBIA"]

            def parse(self, _sheet_name, header=None, nrows=None):
                return preview if header is None else contract_rows

        with patch("lease_accounting.pipeline.processor.pd.ExcelFile", return_value=FakeExcelFile()):
            contracts = self.pipeline._load_contracts("contracts.xlsx")

        self.assertIn("sell_media", contracts.columns)
        self.assertEqual(contracts.iloc[0]["sell_media"], 250000)


if __name__ == "__main__":
    unittest.main()
