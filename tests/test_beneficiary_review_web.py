import unittest
from unittest.mock import patch

from lease_accounting.pipeline.processor import PipelineError
from lease_accounting.web import app


class BeneficiaryReviewWebTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()
        self.item = {
            "review_key": "G112680|900011395",
            "invoice_id": "G112680",
            "source_file": "T025_ARRENDAMIENTO_AGOSTO_2026_1.xml",
            "store": "T025",
            "line_count": 3,
            "items_total": 17324000,
        }
        self.job = {
            "beneficiary_review_context": [self.item],
            "beneficiary_selections": {},
        }

    def test_saves_consolidate_selection_and_continues(self):
        saved_jobs = []
        with (
            patch("lease_accounting.web._load_review_job", return_value=dict(self.job)),
            patch(
                "lease_accounting.web._save_review_job",
                side_effect=lambda _batch_id, job: saved_jobs.append(dict(job)),
            ),
            patch("lease_accounting.web._continue_processing_job", return_value="CONTINUE"),
        ):
            response = self.client.post(
                "/beneficiary-review/test-batch",
                data={
                    "item_count": "1",
                    "review_key_0": self.item["review_key"],
                    "beneficiary_action_0": "consolidate",
                },
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_data(as_text=True), "CONTINUE")
        self.assertEqual(
            saved_jobs[-1]["beneficiary_selections"][self.item["review_key"]],
            "consolidate",
        )

    def test_all_skipped_invoices_render_final_issues_instead_of_redirecting(self):
        with (
            patch("lease_accounting.web._load_review_job", return_value=dict(self.job)),
            patch("lease_accounting.web._save_review_job"),
            patch(
                "lease_accounting.web._continue_processing_job",
                side_effect=PipelineError(
                    "Todas las facturas del lote se marcaron para procesamiento manual. "
                    "No se generaron archivos CSV."
                ),
            ),
            patch("lease_accounting.web._write_processing_issues_report"),
        ):
            response = self.client.post(
                "/beneficiary-review/test-batch",
                data={
                    "item_count": "1",
                    "review_key_0": self.item["review_key"],
                    "beneficiary_action_0": "skip",
                },
            )

        body = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("G112680", body)
        self.assertIn("Facturas omitidas", body)
        self.assertNotIn("Descargas Disponibles", body)


if __name__ == "__main__":
    unittest.main()
