import json
from pathlib import Path
import tempfile
import unittest

from lucy.bundle import InvoiceBundle
from lucy.history import load_pending_validation_history, record_found_invoice


class BotLucyHistoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.bundle_dir = self.root / "bundle"
        self.csv_path = self.bundle_dir / "csv" / "F1.csv"
        self.csv_path.parent.mkdir(parents=True)
        self.csv_path.write_text("account,amount\n1,100\n", encoding="utf-8")
        self.json_path = self.bundle_dir / "control_facturas_bot.json"
        self.json_path.write_text("{}", encoding="utf-8")
        self.history_path = self.root / "persistent-history" / "validation_history.json"
        self.invoice = {
            "folio": "F1",
            "vendor": "V1",
            "partes": [{
                "parte": 1,
                "allocated_costs": "csv/F1.csv",
                "base": 100,
                "iva": 0,
            }],
        }
        self.bundle = InvoiceBundle(
            directory=self.bundle_dir,
            json_path=self.json_path,
            data={"ejecucion": "run-1"},
        )

    def test_found_invoice_and_allocated_costs_persist_until_next_step(self):
        record_found_invoice(
            self.bundle,
            self.invoice,
            "Invoice review complete (not saved)",
            new_attempt=True,
            history_path=self.history_path,
        )

        pending = load_pending_validation_history(self.history_path)

        self.assertEqual(len(pending), 1)
        self.assertTrue(pending[0]["pendiente_siguiente_paso"])
        archived_csv = self.history_path.parent / pending[0]["archivos_csv"][0]["ruta"]
        self.assertEqual(archived_csv.read_text(encoding="utf-8"), "account,amount\n1,100\n")
        self.assertEqual(pending[0]["informacion_factura"]["folio"], "F1")
        self.assertEqual(len(pending[0]["intentos"]), 1)

    def test_found_but_unprocessed_status_updates_without_losing_history(self):
        record_found_invoice(
            self.bundle,
            self.invoice,
            "Invoice review complete (not saved)",
            new_attempt=True,
            history_path=self.history_path,
        )
        record_found_invoice(
            self.bundle,
            self.invoice,
            "found_but_not_processed",
            "TimeoutError: detail did not open",
            history_path=self.history_path,
        )

        pending = load_pending_validation_history(self.history_path)

        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0]["estado_validacion"], "found_but_not_processed")
        self.assertEqual(pending[0]["error"], "TimeoutError: detail did not open")
        self.assertEqual(len(pending[0]["intentos"]), 1)

    def test_history_writer_creates_valid_json(self):
        record_found_invoice(
            self.bundle,
            self.invoice,
            "duplicate_requires_review",
            new_attempt=True,
            history_path=self.history_path,
        )

        history = json.loads(self.history_path.read_text(encoding="utf-8"))
        self.assertEqual(history["schema_version"], 1)
        self.assertEqual(len(history["facturas"]), 1)


if __name__ == "__main__":
    unittest.main()
