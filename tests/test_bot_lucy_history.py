from functools import partial
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from lucy import history
from lucy.bundle import InvoiceBundle
from lucy.history import (
    CLOSED_COMPLETED,
    REASON_NOT_REFLECTED,
    REASON_VALIDATED,
    REASON_VALIDATION_ERROR,
    close_waiting_invoice,
    load_waiting_list,
    record_interactive_export_result,
    upsert_waiting_invoice,
)
from lucy.plan import WorkItem, build_export_items, build_validation_items
from lucy.processes import interactive_export
from lucy.processes.interactive_export import (
    EXPORT_DUPLICATED,
    EXPORT_FOUND,
    EXPORT_NOT_FOUND,
    status_for_export_matches,
    update_waiting_list,
)


class BotLucyWaitingListTests(unittest.TestCase):
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
            "partes": [{"parte": 1, "allocated_costs": "csv/F1.csv", "base": 100, "iva": 0}],
        }
        self.bundle = InvoiceBundle(
            directory=self.bundle_dir,
            json_path=self.json_path,
            data={"ejecucion": "run-1"},
        )

    def upsert(self, reason, status="x", bundle=None, invoice=None):
        upsert_waiting_invoice(
            bundle or self.bundle, invoice or self.invoice, reason, status,
            new_attempt=True, history_path=self.history_path,
        )

    def waiting(self):
        return load_waiting_list(self.history_path)

    def patch_history_path(self):
        names = ["record_interactive_export_result", "close_waiting_invoice", "upsert_waiting_invoice"]
        patches = {
            name: partial(getattr(history, name), history_path=self.history_path) for name in names
        }
        return mock.patch.multiple(interactive_export, **patches)

    def test_validated_invoice_keeps_data_and_csv_in_waiting_list(self):
        self.upsert(REASON_VALIDATED)

        waiting = self.waiting()

        self.assertEqual(len(waiting), 1)
        self.assertEqual(waiting[0]["motivo_espera"], REASON_VALIDATED)
        archived_csv = self.history_path.parent / waiting[0]["archivos_csv"][0]["ruta"]
        self.assertEqual(archived_csv.read_text(encoding="utf-8"), "account,amount\n1,100\n")

    def test_update_from_waiting_list_without_bundle_keeps_archived_csv(self):
        self.upsert(REASON_NOT_REFLECTED)
        upsert_waiting_invoice(
            None, self.invoice, REASON_VALIDATED, "ok", history_path=self.history_path
        )

        waiting = self.waiting()

        self.assertEqual(waiting[0]["motivo_espera"], REASON_VALIDATED)
        self.assertEqual(len(waiting[0]["archivos_csv"]), 1)

    def test_closed_invoice_leaves_waiting_list(self):
        self.upsert(REASON_VALIDATED)

        self.assertTrue(close_waiting_invoice(self.invoice, CLOSED_COMPLETED, history_path=self.history_path))
        self.assertEqual(self.waiting(), [])
        self.assertFalse(close_waiting_invoice(self.invoice, CLOSED_COMPLETED, history_path=self.history_path))
        saved = json.loads(self.history_path.read_text(encoding="utf-8"))
        self.assertEqual(next(iter(saved["facturas"].values()))["estado_final"], CLOSED_COMPLETED)

    def test_legacy_duplicate_records_are_not_pending(self):
        self.history_path.parent.mkdir(parents=True)
        self.history_path.write_text(json.dumps({
            "schema_version": 1,
            "facturas": {"k": {
                "folio": "F9",
                "estado_validacion": "Invoice number duplicated",
                "pendiente_siguiente_paso": True,
                "informacion_factura": {"folio": "F9"},
            }},
        }), encoding="utf-8")

        self.assertEqual(self.waiting(), [])

    def test_export_attempt_ignored_for_invoice_not_in_waiting_list(self):
        self.assertFalse(
            record_interactive_export_result(self.invoice, EXPORT_NOT_FOUND, 0, history_path=self.history_path)
        )
        self.assertFalse(self.history_path.exists())

    def test_validation_plan_skips_already_validated_invoices(self):
        other = {"folio": "F2", "vendor": "V2", "partes": []}
        self.upsert(REASON_VALIDATED)
        upsert_waiting_invoice(
            self.bundle, other, REASON_NOT_REFLECTED, "x", history_path=self.history_path
        )

        validation = build_validation_items(self.bundle, [self.invoice], self.waiting())
        export = build_export_items(self.bundle, [self.invoice], self.waiting())

        self.assertEqual([item.folio for item in validation], ["F2"])
        self.assertIsNone(validation[0].bundle)
        self.assertEqual([item.folio for item in export], ["F1", "F2"])
        self.assertIs(export[0].bundle, self.bundle)

    def test_export_plan_excludes_validation_duplicates(self):
        item = WorkItem(self.invoice, self.bundle, None)

        export = build_export_items(self.bundle, [self.invoice], [], {item.key})

        self.assertEqual(export, [])

    def test_export_status_for_matches(self):
        self.assertEqual(status_for_export_matches(0), EXPORT_NOT_FOUND)
        self.assertEqual(status_for_export_matches(1), EXPORT_FOUND)
        self.assertEqual(status_for_export_matches(2), EXPORT_DUPLICATED)

    def test_not_found_anywhere_is_added_as_not_reflected(self):
        with self.patch_history_path():
            update_waiting_list(WorkItem(self.invoice, self.bundle, None), EXPORT_NOT_FOUND, 0)

        waiting = self.waiting()
        self.assertEqual(waiting[0]["motivo_espera"], REASON_NOT_REFLECTED)
        self.assertEqual(len(waiting[0]["intentos_interactive_export"]), 1)

    def test_validated_invoice_not_yet_in_export_stays_waiting(self):
        self.upsert(REASON_VALIDATED)
        with self.patch_history_path():
            update_waiting_list(WorkItem(self.invoice, None, REASON_VALIDATED), EXPORT_NOT_FOUND, 0)

        waiting = self.waiting()
        self.assertEqual(waiting[0]["motivo_espera"], REASON_VALIDATED)
        self.assertEqual(waiting[0]["estado_interactive_export"], EXPORT_NOT_FOUND)

    def test_found_in_export_completes_waiting_invoice(self):
        self.upsert(REASON_VALIDATION_ERROR)
        with self.patch_history_path():
            update_waiting_list(WorkItem(self.invoice, None, REASON_VALIDATION_ERROR), EXPORT_FOUND, 1)

        self.assertEqual(self.waiting(), [])

    def test_found_in_export_without_waiting_entry_is_not_added(self):
        with self.patch_history_path():
            update_waiting_list(WorkItem(self.invoice, self.bundle, None), EXPORT_FOUND, 1)

        self.assertEqual(self.waiting(), [])

    def test_export_duplicate_is_not_added_to_waiting_list(self):
        with self.patch_history_path():
            update_waiting_list(WorkItem(self.invoice, self.bundle, None), EXPORT_DUPLICATED, 2)

        self.assertEqual(self.waiting(), [])


if __name__ == "__main__":
    unittest.main()
