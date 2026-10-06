import tempfile
import unittest
from pathlib import Path

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class ProrateoMatchingTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.pipeline = LeaseAccountingPipeline(
            Path(self.temp_dir.name),
            account_reference_path=Path(self.temp_dir.name) / "unused.xlsx",
        )

    def merge(self, rules, vendor="289666", store="T104", ceco="8582", control_store=None, macro=None):
        invoices = pd.DataFrame([{
            "invoice_id": "0222675",
            "vendor": vendor,
            "vendor_key": vendor,
            "store": store,
            "store_key": store,
            "ceco": ceco,
            "ceco_key": ceco,
        }])
        control = pd.DataFrame(
            [] if control_store is None else [{
                "vendor_key": vendor, "store_key": control_store,
                "store": control_store, "ceco": ceco or "8582", "vendor_code": vendor,
            }],
            columns=["vendor_key", "store_key", "store", "ceco", "vendor_code"],
        )
        prorateo = pd.DataFrame(rules, columns=[
            "vendor_key", "store_key", "ceco_key", "vw_percent", "vq_percent",
        ])
        prorateo["ceco"] = prorateo["ceco_key"]
        contracts = pd.DataFrame(columns=[
            "ceco_key", "end_of_term", "rent_min", "sell_media", "status_en_rem",
        ])
        return self.pipeline.merge_data({
            "invoices": invoices, "control": control,
            "contracts": contracts, "prorateo": prorateo,
            "macro_database": macro if macro is not None else pd.DataFrame(),
        })

    def test_same_vendor_in_different_stores_uses_specific_rate_regardless_of_order(self):
        rules = [
            ("289666", "R003", "9264", 1.0, 0.0),
            ("289666", "SGH606", "9303", 1.0, 0.0),
            ("289666", "T104", "8582", 0.098, 0.902),
            ("289666", "T115", "8859", 0.056, 0.944),
        ]
        for ordered in [rules, list(reversed(rules))]:
            for store, ceco, vw, vq in [
                ("T104", "8582", 0.098, 0.902),
                ("T115", "8859", 0.056, 0.944),
                ("R003", "9264", 1.0, 0.0),
            ]:
                with self.subTest(store=store, ordered=ordered):
                    merged = self.merge(ordered, store=store, ceco=ceco)
                    self.assertEqual(len(merged), 1)
                    row = merged.iloc[0]
                    self.assertEqual(row["vw_percent"], vw)
                    self.assertEqual(row["vq_percent"], vq)
                    self.assertEqual(row["prorateo_source"], "vendor_store")
                    self.assertEqual(self.pipeline._expense_tax_code(row), "VW" if vw == 1 else "VQ")

    def test_exact_vendor_store_beats_another_vendor_at_same_store(self):
        row = self.merge([
            ("OTHER", "T104", "8582", 1.0, 0.0),
            ("289666", "T104", "8582", 0.098, 0.902),
        ]).iloc[0]
        self.assertEqual(row["vw_percent"], 0.098)
        self.assertEqual(row["vq_percent"], 0.902)

    def test_vendor_ceco_uses_resolved_ceco_when_invoice_ceco_is_missing(self):
        row = self.merge([
            ("289666", "R003", "9264", 1.0, 0.0),
            ("289666", None, "8582", 0.098, 0.902),
        ], ceco=None, control_store="T104").iloc[0]
        self.assertEqual(row["vw_percent"], 0.098)
        self.assertEqual(row["mapped_ceco_key"], "8582")
        self.assertEqual(row["prorateo_source"], "vendor_ceco")

    def test_resolved_store_matches_when_invoice_store_is_missing(self):
        row = self.merge([
            ("289666", "R003", "9264", 1.0, 0.0),
            ("289666", "T104", "8582", 0.098, 0.902),
        ], store=None, control_store="T104").iloc[0]
        self.assertEqual(row["vw_percent"], 0.098)
        self.assertEqual(row["prorateo_source"], "vendor_store")

    def test_store_and_ceco_fallbacks_beat_generic_vendor_rule(self):
        for store, expected_source in [("T104", "store"), ("UNKNOWN", "ceco")]:
            with self.subTest(store=store):
                row = self.merge([
                    ("289666", None, None, 1.0, 0.0),
                    ("OTHER", "T104", "8582", 0.098, 0.902),
                ], store=store).iloc[0]
                self.assertEqual(row["vw_percent"], 0.098)
                self.assertEqual(row["prorateo_source"], expected_source)

    def test_generic_vendor_fallback_is_preserved(self):
        row = self.merge([("289666", None, None, 0.4, 0.6)]).iloc[0]
        self.assertEqual(row["vw_percent"], 0.4)
        self.assertEqual(row["prorateo_source"], "vendor")

    def test_other_location_does_not_supply_vendor_fallback(self):
        row = self.merge([("289666", "R003", "9264", 1.0, 0.0)]).iloc[0]
        self.assertTrue(pd.isna(row["vw_percent"]))
        self.assertTrue(pd.isna(row["vq_percent"]))

    def test_missing_vq_is_derived_from_selected_vw_not_another_rule(self):
        row = self.merge([
            ("289666", None, None, 1.0, 0.0),
            ("289666", "T104", "8582", 0.098, None),
        ]).iloc[0]
        self.assertEqual(row["vw_percent"], 0.098)
        self.assertAlmostEqual(row["vq_percent"], 0.902)

    def test_macro_fallback_does_not_use_another_location_from_same_vendor(self):
        macro = pd.DataFrame([{
            "vendor_key": "289666", "store_key": "T104", "ceco": "8582",
            "profit_center": "8582", "end_of_term": pd.NaT, "status_en_rem": "CESADO",
            "rent_min": 0, "vw_percent": 0.098, "vq_percent": 0.902, "row_rank": 1,
        }])
        row = self.merge(
            [("289666", "R003", "9264", 1.0, 0.0)], macro=macro,
        ).iloc[0]
        self.assertEqual(row["vw_percent"], 0.098)
        self.assertEqual(row["prorateo_source"], "macro")

    def test_selected_t104_rate_preserves_vq_and_adds_equal_vw_reversal(self):
        merged = self.merge([
            ("289666", "R003", "9264", 1.0, 0.0),
            ("289666", "T104", "8582", 0.098, 0.902),
        ])
        self.pipeline.VARIABLE_ACCOUNT = "1537210004"
        self.pipeline.VAT_ACCOUNT = "2408100000"
        merged = merged.assign(
            amount=797801, total=949383, vat=151582,
            invoice_date=pd.Timestamp("2026-09-15"), due_date=pd.NaT,
            withholding_tax=0, payable_rounding=0, payment_terms_text=None,
            item_count_hint=2, ubl_document_type="Invoice", support_source_file=None,
            support_split_factor=1, support_split_index=1, support_flags=None,
            invoice_discounts=lambda frame: [[] for _ in frame.index],
            invoice_line_items=lambda frame: [[
                {"description": "CANON", "amount": 725274},
                {"description": "ADMINISTRACION", "amount": 72527},
            ] for _ in frame.index],
        )
        output, _ = self.pipeline.apply_rules(merged, period="09-2026")
        summary, _, lines = self.pipeline._build_lucy_views(output)
        self.assertEqual(summary.iloc[0]["vat_vw"], 14855)
        self.assertEqual(summary.iloc[0]["vat_vq"], 136727)
        self.assertEqual(lines["tax_code"].tolist(), ["VQ", "VQ", "VW", "VQ", "VQ"])
        self.assertEqual(lines["posting_key"].tolist(), ["40", "40", "40", "50", "50"])
        self.assertEqual(lines["amount"].tolist(), [725274, 72527, 14855, 136727, 14855])
        self.assertEqual(
            lines["line_type"].tolist(),
            ["invoice_item", "invoice_item", "vat_vw", "vat_vq", "vat_vq_reversal"],
        )


if __name__ == "__main__":
    unittest.main()
