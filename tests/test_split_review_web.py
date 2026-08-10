import unittest
from unittest.mock import patch

from lease_accounting.web import _render_split_review, app


class SplitReviewWebTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()
        self.job = {
            "support_split_factors": {
                "C:/uploads/FACTURA-1.xml": 2,
            },
            "support_split_weights": {},
        }

    def test_accepts_custom_percentages_that_sum_to_100(self):
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
                "/split-review/test-batch",
                data={
                    "item_count": "1",
                    "percentage_0_0": "60",
                    "percentage_0_1": "40",
                },
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_data(as_text=True), "CONTINUE")
        self.assertEqual(
            saved_jobs[-1]["support_split_weights"]["C:/uploads/FACTURA-1.xml"],
            ["60", "40"],
        )

    def test_rejects_percentages_that_do_not_sum_to_100(self):
        with (
            patch("lease_accounting.web._load_review_job", return_value=dict(self.job)),
            patch("lease_accounting.web._save_review_job"),
            patch("lease_accounting.web._render_split_review", return_value="REVIEW_AGAIN"),
        ):
            response = self.client.post(
                "/split-review/test-batch",
                data={
                    "item_count": "1",
                    "percentage_0_0": "60",
                    "percentage_0_1": "30",
                },
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_data(as_text=True), "REVIEW_AGAIN")

    def test_review_page_has_cancel_and_return_action(self):
        item = {
            "support_path": "C:/uploads/FACTURA-1.xml",
            "source_file": "FACTURA-1.xml",
            "invoice_id": "FE1",
            "supplier_name": "PROVEEDOR",
            "split_factor": 2,
            "total": 100,
            "equal_percentages": ["50.00", "50.00"],
        }
        with (
            app.test_request_context(),
            patch("lease_accounting.web._build_split_review_items", return_value=[item]),
            patch("lease_accounting.web._save_review_job"),
        ):
            response = _render_split_review("test-batch", dict(self.job))

        self.assertIn('href="/"', response)
        self.assertIn("Cancelar y volver", response)


if __name__ == "__main__":
    unittest.main()
