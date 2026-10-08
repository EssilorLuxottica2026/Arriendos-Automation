import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class GcContractAccountTests(unittest.TestCase):
    def setUp(self):
        self.fixed_account = "9000000001"
        self.variable_account = "9000000002"
        self.gc_variable_account = "9000000003"
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        workbook_path = Path(self.temp_dir.name) / "FACTURAS_CONTABILIZADAS.xlsx"
        pd.DataFrame(
            [
                {"CUENTA CONTABLE": self.fixed_account, "NOMBRE CUENTA": "ARRIENDOS FIJOS", "SERIE": "RF"},
                {"CUENTA CONTABLE": self.variable_account, "NOMBRE CUENTA": "ARRIENDO VARIABLE", "SERIE": "RV"},
                {"CUENTA CONTABLE": self.fixed_account, "NOMBRE CUENTA": "ADMINISTRACION SELL CAM", "SERIE": "GC"},
                {"CUENTA CONTABLE": self.gc_variable_account, "NOMBRE CUENTA": "ADMINISTRACION VARIABLE", "SERIE": "GC"},
                {"CUENTA CONTABLE": self.gc_variable_account, "NOMBRE CUENTA": "ADMINISTRACION VARIABLE", "SERIE": "GC VARIABLE"},
            ]
        ).to_excel(workbook_path, sheet_name="CONSOLIDADO DE CUENTAS ARRI", index=False)
        self.pipeline = LeaseAccountingPipeline("scratch", account_reference_path=workbook_path)
        self.item = {"description": "CUOTA DE ADMINISTRACION", "amount": 1000}
        self.pipeline.learning_dictionary = pd.DataFrame(
            [
                {
                    "concept": "GC",
                    "account": self.gc_variable_account,
                    "phrases": "CUOTA DE ADMINISTRACION",
                    "text": "GC",
                }
            ]
        )

    def test_positive_sell_media_uses_fixed_gc_account(self):
        concept, account = self.pipeline._item_line_classification(
            self.item,
            {"account": self.gc_variable_account, "rent_type": "RV", "sell_media": 125000},
        )

        self.assertEqual(concept, "GC")
        self.assertEqual(account, self.fixed_account)

    def test_user_learned_gc_selection_still_uses_contract_account(self):
        self.pipeline.learning_dictionary = pd.DataFrame(
            [
                {
                    "concept": "GC",
                    "account": self.gc_variable_account,
                    "phrases": "ADM GC",
                    "text": "GC",
                }
            ]
        )
        concept, account = self.pipeline._item_line_classification(
            {"description": "ADM GC % 01.06-30.06.2026", "amount": 165196},
            {"account": self.gc_variable_account, "rent_type": "RV", "sell_media": 1545851},
        )

        self.assertEqual(concept, "GC")
        self.assertEqual(account, self.fixed_account)

    def test_blank_dash_zero_or_missing_sell_media_use_variable_gc(self):
        for sell_media in [None, "-", "", 0]:
            with self.subTest(sell_media=sell_media):
                concept, account = self.pipeline._item_line_classification(
                    self.item,
                    {"account": self.fixed_account, "rent_type": "RF", "sell_media": sell_media},
                )
                self.assertEqual(concept, "GC V")
                self.assertEqual(account, self.gc_variable_account)

    def test_gc_v_uses_sell_cam_for_administration_account(self):
        self.pipeline.learning_dictionary.at[0, "concept"] = "GC V"
        self.pipeline.learning_dictionary.at[0, "text"] = "GC V"
        for sell_cam, expected in [
            (0, ("GV", self.gc_variable_account)),
            (1000, ("GF", self.fixed_account)),
        ]:
            with self.subTest(sell_cam=sell_cam):
                self.assertEqual(
                    self.pipeline._item_line_classification(
                        self.item, {"sell_media": 1000, "sell_cam": sell_cam},
                    ),
                    expected,
                )

    def test_shared_invoice_phrase_is_classified_as_dynamic_gc_v(self):
        phrase = "SERV PARA USO Y APROVEC DEL INMUEB LOCALES"
        self.pipeline.learning_dictionary = pd.DataFrame([
            {"concept": "RENTA", "account": "SEGUN_CONTRATO", "phrases": "OTRA FRASE", "text": "RENTA"},
            {"concept": "GC V", "account": self.gc_variable_account, "phrases": phrase, "text": "GC V"},
        ])
        concept, account = self.pipeline._item_line_classification(
            {"description": f"{phrase} AGOSTO 2026", "amount": 100},
            {"sell_media": 0},
        )
        self.assertEqual(concept, "GV")
        self.assertEqual(account, self.gc_variable_account)

    def test_sell_cam_classification_emits_negative_credit_note_line_with_key_50(self):
        phrase = "SERV PARA USO Y APROVEC DEL INMUEB LOCALES"
        self.pipeline.learning_dictionary = pd.DataFrame(
            [{"concept": "GC V", "account": self.gc_variable_account, "phrases": phrase, "text": "GC V"}]
        )
        output_df = pd.DataFrame(
            [
                {
                    "invoice_id": "CN-1",
                    "vendor": "100",
                    "store": "T001",
                    "ceco": "6001",
                    "profit_center": "6001",
                    "account": self.fixed_account,
                    "concept": "RF",
                    "rent_type": "RF",
                    "sell_media": 1000,
                    "sell_cam": 1000,
                    "amount": 100,
                    "gross_amount": 100,
                    "discount_total": 0,
                    "posting_discount_total": 0,
                    "payable_rounding": 0,
                    "invoice_total": 100,
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
                    "ubl_document_type": "CreditNote",
                    "invoice_line_items": [{"description": phrase, "amount": 100}],
                    "invoice_discounts": [],
                    "support_split_index": 1,
                    "support_split_factor": 1,
                }
            ]
        )

        _, _, lines = self.pipeline._build_lucy_views(output_df)
        item_line = lines.loc[lines["line_type"].eq("invoice_item")].iloc[0]

        self.assertEqual(item_line["concept"], "GF")
        self.assertEqual(item_line["account"], self.fixed_account)
        self.assertEqual(item_line["posting_key"], "50")
        self.assertEqual(item_line["amount"], -100)

    def test_fixed_space_phrase_uses_contract_rent_status(self):
        phrase = "CONCESION DE ESPACIO FIJO LOCALES"
        self.pipeline.learning_dictionary = pd.DataFrame([
            {"concept": "RENTA", "account": "SEGUN_CONTRATO", "phrases": phrase, "text": "RENTA"},
            {"concept": "GC", "account": self.gc_variable_account, "phrases": "OTRA FRASE", "text": "GC"},
        ])
        concept, account = self.pipeline._item_line_classification(
            {"description": f"{phrase} AGOSTO 2026", "amount": 100},
            {"account": "1537210004", "rent_type": "RV", "sell_media": 0},
        )
        self.assertEqual(concept, "RV")
        self.assertEqual(account, "1537210004")

    def test_shared_account_preserves_distinct_abbreviations(self):
        self.pipeline.learning_dictionary = pd.DataFrame([
            {"concept": "AIRE", "account": "1532001600", "phrases": "AIRE ACONDICIONADO", "text": "AIRE"},
            {"concept": "ASEO", "account": "1532001600", "phrases": "REINTEGRO DE ASEO", "text": "ASEO"},
            {"concept": "VIGILANCIA", "account": "1532001600", "phrases": "VIGILANCIA", "text": "VIGILANCIA"},
            {"concept": "FP", "account": "1537210003", "phrases": "FONDO DE PROMOCION", "text": "FP"},
            {"concept": "GC OFICINA", "account": "1535100020", "phrases": "COMMON AREA MAINTENANCE", "text": "GC OFICINA"},
            {"concept": "RF OFICINA", "account": "1535100000", "phrases": "ARRIENDOS INMUEBLES", "text": "RF OFICINA"},
        ])
        for _, rule in self.pipeline.learning_dictionary.iterrows():
            with self.subTest(concept=rule["concept"]):
                self.assertEqual(
                    self.pipeline._item_line_classification(
                        {"description": rule["phrases"], "amount": 100},
                        {"sell_media": 0},
                    ),
                    (rule["text"], rule["account"]),
                )

    def test_allocated_cost_text_uses_gc_variable(self):
        output_df = pd.DataFrame(
            [
                {
                    "invoice_id": "GC1",
                    "vendor": "100",
                    "store": "T001",
                    "ceco": "6001",
                    "profit_center": "6001",
                    "account": self.fixed_account,
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

        self.assertEqual(item_line["concept"], "GC V")
        self.assertEqual(item_line["account"], self.gc_variable_account)
        self.assertIn("GC V", item_line["text"])

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
                    "SELL CAM": 175000,
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
        self.assertIn("sell_cam", contracts.columns)
        self.assertEqual(contracts.iloc[0]["sell_cam"], 175000)


if __name__ == "__main__":
    unittest.main()
