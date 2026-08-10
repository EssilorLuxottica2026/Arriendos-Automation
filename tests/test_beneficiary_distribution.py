import unittest
from pathlib import Path

import pandas as pd

from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader
from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class BeneficiaryDistributionTests(unittest.TestCase):
    def setUp(self):
        self.fixture = Path(__file__).parent / "fixtures" / "invoice_beneficiary_distribution.xml"
        self.parsed = InvoiceSupportReader().parse_file(self.fixture)
        self.pipeline = object.__new__(LeaseAccountingPipeline)
        self.pipeline.warnings = []
        self.pipeline.logs = []
        self.pipeline.preflight_processing_issues = []
        self.pipeline.beneficiary_review_items = []

    def _row(self, line_items=None):
        return pd.Series(
            {
                "invoice_id": "BEN-1",
                "invoice_nit_key": "BEN1|816003192",
                "support_source_file": self.fixture.name,
                "store": "T001",
                "ceco": "1001",
                "support_subtotal": 300,
                "total": 357,
                "vat": 57,
                "amount": 300,
                "support_split_factor": 3,
                "invoice_discounts": [],
                "invoice_line_items": line_items if line_items is not None else self.parsed.line_items,
            }
        )

    def test_parser_preserves_distinct_line_beneficiaries(self):
        self.assertEqual([item["beneficiary_id"] for item in self.parsed.line_items], ["111", "222", "333"])
        self.assertEqual({item["item_code"] for item in self.parsed.line_items}, {"CANON"})
        self.assertEqual({item["tax_percent"] for item in self.parsed.line_items}, {19.0})

    def test_detects_only_identical_lines_with_distinct_beneficiaries(self):
        candidate = self.pipeline._beneficiary_distribution_candidate(self._row())

        self.assertEqual(candidate["beneficiary_count"], 3)
        self.assertEqual(candidate["items_total"], 300)

        different_items = [dict(item) for item in self.parsed.line_items]
        different_items[1]["description"] = "ADMINISTRACION"
        self.assertIsNone(
            self.pipeline._beneficiary_distribution_candidate(self._row(different_items))
        )

    def test_consolidates_before_split_only_when_selected(self):
        invoices = pd.DataFrame([self._row()])

        result = self.pipeline._apply_beneficiary_distribution_selections(
            invoices,
            {"BEN1|816003192": "consolidate"},
        )

        items = result.iloc[0]["invoice_line_items"]
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["amount"], 300)
        self.assertEqual(items[0]["beneficiary_ids"], ["111", "222", "333"])

    def test_skip_adds_final_processing_issue(self):
        invoices = pd.DataFrame([self._row(), self._row()])
        invoices.at[1, "invoice_id"] = "NORMAL-1"
        invoices.at[1, "invoice_nit_key"] = "NORMAL1|816003192"
        normal_items = [dict(self.parsed.line_items[0])]
        normal_items[0]["amount"] = 50
        invoices.at[1, "invoice_line_items"] = normal_items

        result = self.pipeline._apply_beneficiary_distribution_selections(
            invoices,
            {"BEN1|816003192": "skip"},
        )

        self.assertEqual(result["invoice_id"].tolist(), ["NORMAL-1"])
        self.assertEqual(
            self.pipeline.preflight_processing_issues[0]["issue_type"],
            "beneficiary_distribution_skipped",
        )


if __name__ == "__main__":
    unittest.main()
