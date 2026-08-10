import unittest

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class ProcessingIssueSolutionTests(unittest.TestCase):
    def setUp(self):
        self.pipeline = object.__new__(LeaseAccountingPipeline)

    def test_builds_actionable_solutions_for_each_validation_problem(self):
        validation = pd.DataFrame(
            [
                {
                    "invoice_id": "FE1",
                    "source_file": "FE1.xml",
                    "store": "T001",
                    "ceco": "1001",
                    "missing_ceco": False,
                    "missing_contract": True,
                    "missing_prorrateo": True,
                    "missing_amount": False,
                    "blocked_replacement_status": False,
                    "unresolved_discount_note": False,
                    "invalid_discount": False,
                }
            ]
        )

        issue = self.pipeline._build_processing_issues(validation)[0]

        self.assertIn("Contratos_con_condiciones", issue["possible_solution"])
        self.assertIn("PRORATEO", issue["possible_solution"])


if __name__ == "__main__":
    unittest.main()
