import unittest

import pandas as pd

from src.lease_accounting.pipeline.processor import LeaseAccountingPipeline


class ExpenseTaxCodeTests(unittest.TestCase):
    def setUp(self):
        self.pipeline = LeaseAccountingPipeline("scratch")

    def test_expense_tax_code_uses_full_vw_prorate_without_vat(self):
        for vat_total in (None, 0):
            with self.subTest(vat_total=vat_total):
                self.assertEqual(
                    self.pipeline._expense_tax_code(
                        {"vat_total": vat_total, "vw_percent": 1}
                    ),
                    "VW",
                )

    def test_expense_tax_code_uses_vq_when_prorate_is_not_full_vw(self):
        for vat_total in (None, 0):
            with self.subTest(vat_total=vat_total):
                self.assertEqual(
                    self.pipeline._expense_tax_code(
                        {"vat_total": vat_total, "vw_percent": 0.4}
                    ),
                    "VQ",
                )

    def test_expense_tax_code_defaults_to_vq_when_vw_prorate_is_missing(self):
        self.assertEqual(
            self.pipeline._expense_tax_code({"vat_total": 0}),
            "VQ",
        )

if __name__ == "__main__":
    unittest.main()
