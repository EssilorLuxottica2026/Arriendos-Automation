import unittest

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class SplitPercentageTests(unittest.TestCase):
    def test_custom_60_40_split_reconciles_every_amount(self):
        pipeline = LeaseAccountingPipeline("scratch")
        invoices = pd.DataFrame(
            [
                {
                    "invoice_id": "PRUEBA-6040",
                    "support_split_factor": 2,
                    "support_split_weights": ["60", "40"],
                    "total": 1001,
                    "vat": 191,
                    "amount": 1001,
                    "support_subtotal": 810,
                    "support_total": 1001,
                    "support_detected_iva": 191,
                    "invoice_line_items": [
                        {
                            "description": "CUOTA DE ADMINISTRACION",
                            "amount": 1001,
                            "net_amount": 1001,
                            "tax_amount": 191,
                        }
                    ],
                    "invoice_discounts": [
                        {
                            "description": "DESCUENTO",
                            "amount": 101,
                            "base_amount": 1001,
                        }
                    ],
                }
            ]
        )

        result = pipeline._expand_split_invoice_rows(invoices)

        self.assertEqual(result["total"].tolist(), [601, 400])
        self.assertEqual(result["vat"].tolist(), [115, 76])
        self.assertEqual(result["support_subtotal"].tolist(), [486, 324])
        self.assertEqual(result["support_split_percent"].tolist(), [60.0, 40.0])
        self.assertEqual(
            [row[0]["amount"] for row in result["invoice_line_items"]],
            [601, 400],
        )
        self.assertEqual(
            [row[0]["amount"] for row in result["invoice_discounts"]],
            [61, 40],
        )
        self.assertEqual(result["total"].sum(), 1001)
        self.assertEqual(result["vat"].sum(), 191)

    def test_three_way_split_assigns_rounding_residual_to_last_file(self):
        pipeline = LeaseAccountingPipeline("scratch")
        invoices = pd.DataFrame(
            [
                {
                    "invoice_id": "PRUEBA-3",
                    "support_split_factor": 3,
                    "support_split_weights": ["33.33", "33.33", "33.34"],
                    "total": 100,
                    "amount": 100,
                    "invoice_line_items": [{"description": "CANON", "amount": 100}],
                    "invoice_discounts": [],
                }
            ]
        )

        result = pipeline._expand_split_invoice_rows(invoices)

        self.assertEqual(result["total"].tolist(), [33, 33, 34])
        self.assertEqual(
            [row[0]["amount"] for row in result["invoice_line_items"]],
            [33, 33, 34],
        )
        self.assertEqual(result["total"].sum(), 100)


if __name__ == "__main__":
    unittest.main()
