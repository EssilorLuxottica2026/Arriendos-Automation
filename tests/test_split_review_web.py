import unittest
from unittest.mock import patch

from lease_accounting.web import app


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


if __name__ == "__main__":
    unittest.main()
