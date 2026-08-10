import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from lease_accounting.web import _write_processing_issues_report


class ProcessingIssuesReportTests(unittest.TestCase):
    def test_writes_downloadable_excel_with_reason_and_solution(self):
        output_dir = Path("scratch") / "processing_issue_report_test"
        (output_dir / "run-test").mkdir(parents=True, exist_ok=True)
        result = {
            "output_csv": "run-test/allocated.zip",
            "warnings": [],
            "processing_issues": [
                {
                    "invoice_id": "FE1",
                    "source_file": "FE1.xml",
                    "store": "T001",
                    "problem": "El NIT no coincide",
                    "possible_solution": "Corrige el NIT en el invoices file",
                }
            ],
        }

        with patch("lease_accounting.web.OUTPUT_DIR", output_dir):
            _write_processing_issues_report(result)

        report_path = output_dir / result["omitted_invoices_report"]
        workbook = load_workbook(report_path, read_only=True)
        sheet = workbook["Facturas omitidas"]
        self.assertEqual(sheet["A2"].value, "FE1")
        self.assertEqual(sheet["D2"].value, "El NIT no coincide")
        self.assertEqual(sheet["E2"].value, "Corrige el NIT en el invoices file")


if __name__ == "__main__":
    unittest.main()
