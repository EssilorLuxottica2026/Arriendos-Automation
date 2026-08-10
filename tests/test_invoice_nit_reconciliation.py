import unittest

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class InvoiceNitReconciliationTests(unittest.TestCase):
    def setUp(self):
        self.pipeline = object.__new__(LeaseAccountingPipeline)
        self.pipeline.warnings = []
        self.pipeline.logs = []

    def _invoices(self):
        rows = [
            ("U1", "111", "T139", "Q137"),
            ("U2", "222", "T139", "Q137"),
            ("U3", "333", "T139", "Q137"),
        ]
        return pd.DataFrame(
            [
                {
                    "invoice_pdf_key": invoice_id,
                    "invoice_nit_key": f"{invoice_id}|{nit}",
                    "supplier_nit": nit,
                    "store_key": store,
                    "ceco_key": ceco,
                    "store": store,
                    "ceco": ceco,
                }
                for invoice_id, nit, store, ceco in rows
            ]
        )

    def _supports(self, nits):
        return pd.DataFrame(
            [
                {
                    "invoice_id": invoice_id,
                    "invoice_key": invoice_id,
                    "supplier_nit": nit,
                    "invoice_nit_key": f"{invoice_id}|{nit}",
                }
                for invoice_id, nit in zip(("U1", "U2", "U3"), nits)
            ]
        )

    def test_reconciles_complete_nit_permutation_in_same_store_and_ceco(self):
        result = self.pipeline._reconcile_permuted_support_nits(
            self._invoices(),
            self._supports(("222", "333", "111")),
        )

        self.assertEqual(result["invoice_nit_key"].tolist(), ["U1|111", "U2|222", "U3|333"])
        self.assertIn("NIT intercambiados", self.pipeline.warnings[0])
        self.assertIn("T139 / CeCo Q137", self.pipeline.warnings[0])

    def test_does_not_reconcile_incomplete_or_unknown_nit_set(self):
        result = self.pipeline._reconcile_permuted_support_nits(
            self._invoices(),
            self._supports(("222", "333", "999")),
        )

        self.assertEqual(result["invoice_nit_key"].tolist(), ["U1|222", "U2|333", "U3|999"])
        self.assertEqual(self.pipeline.warnings, [])


if __name__ == "__main__":
    unittest.main()
