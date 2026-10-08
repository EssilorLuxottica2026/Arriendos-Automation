import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class OfficeAccountMappingTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.pipeline = LeaseAccountingPipeline(
            Path(self.temp_dir.name),
            account_reference_path=Path(self.temp_dir.name) / "unused.xlsx",
        )
        self.pipeline.learning_dictionary = pd.DataFrame(
            [
                {"concept": "RENTA", "account": "SEGUN_CONTRATO", "phrases": "ARRIENDAMIENTO INMUEBLE", "text": "RENTA"},
                {"concept": "RF OFICINA", "account": "1535100000", "phrases": "ARRIENDOS INMUEBLES", "text": "RENTA"},
                {"concept": "ENERGIA", "account": "1532001605", "phrases": "MANDATO EPM ENERGIA", "text": "ENERGIA"},
                {"concept": "ENERGIA OFICINA", "account": "1537180001", "phrases": "MANDATO EPM ENERGIA", "text": "ENERGIA"},
                {"concept": "ACUED", "account": "1532001604", "phrases": "ACUEDUCTO", "text": "ACUED"},
                {"concept": "ACUED OFICINA", "account": "1534001605", "phrases": "AGUA Y DRENAJE OFICINA", "text": "ACUED"},
                {"concept": "GC V", "account": "1537210002", "phrases": "ADMINISTRACION VARIABLE", "text": "GC V"},
                {"concept": "GC OFICINA", "account": "1535100020", "phrases": "ADMINISTRACION VARIABLE OFICINA", "text": "GC V"},
            ]
        )

    def test_energy_uses_office_account_only_for_office_store(self):
        description = "MANDATO EPM ENERGIA"
        office = self.pipeline._item_line_classification(
            {"description": description, "amount": 100},
            {"store": "OFICINA BOGOTA"},
        )
        regular = self.pipeline._item_line_classification(
            {"description": description, "amount": 100},
            {"store": "T001"},
        )

        self.assertEqual(office, ("ENERGIA", "1537180001"))
        self.assertEqual(regular, ("ENERGIA", "1532001605"))

    def test_office_water_and_administration_keep_common_siglas(self):
        water = self.pipeline._item_line_classification(
            {"description": "AGUA Y DRENAJE OFICINA", "amount": 100},
            {"store": "OFICINA"},
        )
        administration = self.pipeline._item_line_classification(
            {"description": "ADMINISTRACION VARIABLE OFICINA", "amount": 100},
            {"store": "OFICINA", "sell_cam": 0, "sell_media": 0},
        )

        self.assertEqual(water, ("ACUED", "1534001605"))
        self.assertEqual(administration, ("GV", "1535100020"))

    def test_office_rent_uses_one_account_and_keeps_contract_sigla(self):
        for rent_type in ["RF", "RV"]:
            with self.subTest(rent_type=rent_type):
                result = self.pipeline._item_line_classification(
                    {"description": "ARRIENDAMIENTO INMUEBLE", "amount": 100},
                    {
                        "store": "OFICINA",
                        "account": "1245150017",
                        "rent_type": rent_type,
                    },
                )
                self.assertEqual(result, (rent_type, "1535100000"))

    def test_office_only_dictionary_entry_does_not_match_regular_store(self):
        result = self.pipeline._item_line_classification(
            {"description": "AGUA Y DRENAJE OFICINA", "amount": 100},
            {"store": "T001"},
        )

        self.assertEqual(result, ("SIN CLASIFICAR", ""))

    def test_summary_rent_uses_office_account_and_preserves_rf(self):
        for status, expected_concept in [("VIGENTE", "RF"), ("CESADO", "RV")]:
            with self.subTest(status=status):
                merged = pd.DataFrame(
                    [
                        {
                            "invoice_id": "OFFICE-1",
                            "vendor": "100",
                            "store": "OFICINA",
                            "mapped_store": "OFICINA",
                            "ceco": "6001",
                            "mapped_ceco": "6001",
                            "profit_center": "6001",
                            "status_en_rem": status,
                            "amount": 1000,
                            "total": 1190,
                            "vat": 190,
                            "invoice_date": pd.Timestamp("2026-08-01"),
                            "due_date": pd.NaT,
                            "withholding_tax": 0,
                            "payable_rounding": 0,
                            "payment_terms_text": None,
                            "item_count_hint": 1,
                            "ubl_document_type": "Invoice",
                            "support_source_file": None,
                            "support_split_factor": 1,
                            "support_split_index": 1,
                            "support_flags": None,
                            "invoice_discounts": [],
                            "invoice_line_items": [],
                            "vw_percent": 1.0,
                            "vq_percent": 0.0,
                            "sell_media": 0,
                            "sell_cam": 0,
                        }
                    ]
                )
                with patch.object(self.pipeline, "_apply_distribution", side_effect=lambda frame: frame):
                    output, _ = self.pipeline.apply_rules(merged, period="08-2026")

                self.assertEqual(output.iloc[0]["account"], "1535100000")
                self.assertEqual(output.iloc[0]["concept"], expected_concept)


if __name__ == "__main__":
    unittest.main()
