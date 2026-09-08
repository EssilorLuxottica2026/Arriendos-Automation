from io import BytesIO
import unittest
from unittest.mock import MagicMock, patch

from lease_accounting import web


class MasterUploadFlowTests(unittest.TestCase):
    def test_index_opens_when_no_master_databases_exist(self):
        status = [
            {"filename": "CONTROL_ARRI_ADMON.xlsx", "required": True, "available": False},
        ]
        with patch.object(web, "_master_status", return_value=status):
            with web.app.test_client() as client:
                response = client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b'data-file-status="control_file"', response.data)
        self.assertIn(b'data-initial-available="false"', response.data)

    def test_processing_prompts_for_all_missing_required_databases(self):
        missing_input_dir = MagicMock()
        missing_input_dir.__truediv__.return_value.exists.return_value = False
        dictionary_path = MagicMock()
        dictionary_path.exists.return_value = True
        with (
            patch.object(web, "INPUT_DIR", missing_input_dir),
            patch.object(web, "LEARNING_DICTIONARY_PATH", dictionary_path),
            web.app.test_client() as client,
        ):
            response = client.post(
                "/process",
                data={"invoices_file": (BytesIO(b"Invoice_ID\nINV-1\n"), "invoices.csv")},
                content_type="multipart/form-data",
                follow_redirects=True,
            )

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"CONTROL_ARRI_ADMON.xlsx", response.data)
        self.assertIn(b"Contratos_con_condiciones.xlsx", response.data)

    def test_save_masters_does_not_require_invoices_file(self):
        with patch.object(web, "_resolve_and_save_master", return_value=None) as save_master:
            with web.app.test_client() as client:
                response = client.post(
                    "/save-masters",
                    data={
                        "control_file": (BytesIO(b"control"), "control.xlsx"),
                    },
                    content_type="multipart/form-data",
                )

        self.assertEqual(response.status_code, 302)
        save_master.assert_called_once()
        self.assertEqual(save_master.call_args.args[0], "control_file")


if __name__ == "__main__":
    unittest.main()
