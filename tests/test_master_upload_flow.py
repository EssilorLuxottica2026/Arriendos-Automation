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
        self.assertIn(b"pendiente de cargar", response.data)

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


if __name__ == "__main__":
    unittest.main()
