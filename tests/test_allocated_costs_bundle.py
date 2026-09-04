import csv
from pathlib import Path
import tempfile
import unittest
import zipfile

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class AllocatedCostsBundleTests(unittest.TestCase):
    def test_writes_one_isolated_csv_per_invoice(self):
        test_root = Path.cwd() / "scratch"
        test_root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=test_root) as temp_dir:
            output_dir = Path(temp_dir)
            pipeline = LeaseAccountingPipeline(output_dir)
            summary_df = pd.DataFrame(
                [
                    {
                        "invoice_id": "FACTURA-1",
                        "ceco": "1001",
                        "profit_center": "1001",
                        "posting_count": 2,
                        "posting_index": 1,
                    },
                    {
                        "invoice_id": "FACTURA-1",
                        "ceco": "1001",
                        "profit_center": "1001",
                        "posting_count": 2,
                        "posting_index": 2,
                    },
                    {
                        "invoice_id": "FACTURA-2",
                        "ceco": "2002",
                        "profit_center": "2002",
                        "posting_count": 1,
                        "posting_index": 1,
                    },
                ]
            )
            lucy_export_df = pd.DataFrame(
                [
                    {
                        "Attribuzione$Assignment": "FACTURA-1",
                        "CdC$Cost Center": "1001",
                        "ProfitCenter$Profit Center": "1001",
                        "Imp$Amount": 125,
                        "_posting_index": 1,
                    },
                    {
                        "Attribuzione$Assignment": "FACTURA-1",
                        "CdC$Cost Center": "1001",
                        "ProfitCenter$Profit Center": "1001",
                        "Imp$Amount": 125,
                        "_posting_index": 2,
                    },
                    {
                        "Attribuzione$Assignment": "FACTURA-2",
                        "CdC$Cost Center": "2002",
                        "ProfitCenter$Profit Center": "2002",
                        "Imp$Amount": 275,
                        "_posting_index": 1,
                    },
                ]
            )

            pipeline._write_allocated_costs_csv_bundle(
                summary_df=summary_df,
                lucy_export_df=lucy_export_df,
                bundle_name="allocated_costs.zip",
                timestamp="test",
            )

            with zipfile.ZipFile(output_dir / "allocated_costs.zip") as bundle:
                self.assertEqual(
                    sorted(bundle.namelist()),
                    [
                        "FACTURA-1_1_allocated_costs.csv",
                        "FACTURA-1_2_allocated_costs.csv",
                        "FACTURA-2_allocated_costs.csv",
                    ],
                )
                first_rows = list(
                    csv.DictReader(
                        bundle.read("FACTURA-1_1_allocated_costs.csv")
                        .decode("utf-8-sig")
                        .splitlines()
                    )
                )
                first_split_rows = list(
                    csv.DictReader(
                        bundle.read("FACTURA-1_2_allocated_costs.csv")
                        .decode("utf-8-sig")
                        .splitlines()
                    )
                )
                second_rows = list(
                    csv.DictReader(
                        bundle.read("FACTURA-2_allocated_costs.csv")
                        .decode("utf-8-sig")
                        .splitlines()
                    )
                )

            self.assertEqual([row["Attribuzione$Assignment"] for row in first_rows], ["FACTURA-1"])
            self.assertEqual(
                [row["Attribuzione$Assignment"] for row in first_split_rows],
                ["FACTURA-1"],
            )
            self.assertEqual([row["Attribuzione$Assignment"] for row in second_rows], ["FACTURA-2"])
            self.assertEqual([row["Imp$Amount"] for row in first_rows], ["125"])
            self.assertEqual([row["Imp$Amount"] for row in first_split_rows], ["125"])
            self.assertEqual([row["CdC$Cost Center"] for row in first_rows], [""])
            self.assertEqual([row["CdC$Cost Center"] for row in first_split_rows], [""])
            self.assertEqual([row["CdC$Cost Center"] for row in second_rows], [""])
            self.assertNotIn("_posting_index", first_rows[0])
            self.assertNotIn("_posting_index", first_split_rows[0])
            self.assertNotIn("_posting_index", second_rows[0])

    def test_lucy_export_keeps_cost_center_blank_and_profit_center_filled(self):
        pipeline = LeaseAccountingPipeline("scratch")
        lines_df = pd.DataFrame(
            [
                {
                    "posting_index": 1,
                    "posting_key": "40",
                    "amount": 1000,
                    "tax_code": "V0",
                    "account": "1537210002",
                    "ceco": "6001",
                    "profit_center": "6001",
                    "text": "08-2026 GC VARIABLE T001 6001",
                    "reference": "F1",
                    "store": "T001",
                }
            ]
        )

        export = pipeline._build_macro_lucy_export(
            pd.DataFrame(),
            lines_df,
            pd.DataFrame(),
        )

        self.assertTrue(export["CdC$Cost Center"].isna().all())
        self.assertEqual(export.iloc[0]["ProfitCenter$Profit Center"], "6001")

    def test_sell_cam_line_payload_uses_fixed_account_when_value_exists(self):
        pipeline = LeaseAccountingPipeline(output_dir=Path("tmp_output"))

        row = {
            "sell_cam": 50000,
            "account": "1537210004",
            "store": "T1",
            "ceco": "CCO1",
            "invoice_date": pd.Timestamp("2026-01-10"),
        }

        account, text, amount = pipeline._sell_cam_line_payload(row)

        self.assertEqual(amount, 50000)
        self.assertEqual(account, pipeline.FIXED_ACCOUNT)
        self.assertEqual(text, "Administracion fijo")

    def test_sell_cam_line_payload_uses_variable_account_when_empty(self):
        pipeline = LeaseAccountingPipeline(output_dir=Path("tmp_output"))

        row = {
            "sell_cam": "",
            "account": "1245150017",
            "store": "T1",
            "ceco": "CCO1",
            "invoice_date": pd.Timestamp("2026-01-10"),
        }

        account, text, amount = pipeline._sell_cam_line_payload(row)

        self.assertEqual(amount, 0)
        self.assertEqual(account, pipeline.VARIABLE_ACCOUNT)
        self.assertEqual(text, "Administracion variado")

if __name__ == "__main__":
    unittest.main()
