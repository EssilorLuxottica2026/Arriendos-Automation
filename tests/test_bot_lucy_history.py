from functools import partial
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from lucy import history
from lucy.bundle import InvoiceBundle, allocated_costs_relative_path
from lucy.history import (
    CLOSED_COMPLETED,
    REASON_EXPORT_ERROR,
    REASON_NOT_REFLECTED,
    REASON_VALIDATED,
    REASON_VALIDATION_ERROR,
    close_waiting_invoice,
    load_waiting_list,
    record_interactive_export_result,
    upsert_waiting_invoice,
)
from lucy.notifications import error_messages
from lucy.plan import WorkItem, build_export_items, build_validation_items
from lucy.processes import export_document, interactive_export
from lucy.processes.export_document import (
    VAT_CODE_WITH_VAT,
    VAT_CODE_WITHOUT_VAT,
    ZERO_BALANCE,
    allocated_costs_csv,
    vat_code_for,
)
from lucy.processes.interactive_export import (
    EXPORT_DUPLICATED,
    EXPORT_FOUND,
    EXPORT_NOT_COMPLETED,
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
            "partes": [{
                "parte": 1,
                "csv": {"allocated_costs": "csv\\F1.csv"},
                "allocated_costs": [{"Imp$Amount": 100, "Cod_Iva$Vat Code": "VQ"}],
                "base": 100,
                "iva": 0,
            }],
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

    def test_failed_completion_in_export_keeps_invoice_waiting(self):
        with self.patch_history_path():
            update_waiting_list(WorkItem(self.invoice, self.bundle, None), EXPORT_NOT_COMPLETED, 1, "boom")

        waiting = self.waiting()
        self.assertEqual(waiting[0]["motivo_espera"], REASON_EXPORT_ERROR)
        self.assertEqual(waiting[0]["error"], "boom")

    def test_csv_path_comes_from_csv_section_not_from_rows(self):
        self.assertEqual(allocated_costs_relative_path(self.invoice["partes"][0]), "csv/F1.csv")
        self.assertIsNone(allocated_costs_relative_path({"allocated_costs": [{"Imp$Amount": 1}]}))

    def test_export_csv_from_run_bundle(self):
        path = allocated_costs_csv(WorkItem(self.invoice, self.bundle, None))

        self.assertEqual(path, self.csv_path.resolve())

    def test_export_csv_from_waiting_list_archive(self):
        self.upsert(REASON_VALIDATED)
        archived = partial(history.archived_allocated_costs_csv, history_path=self.history_path)

        with mock.patch.object(export_document, "archived_allocated_costs_csv", archived):
            path = allocated_costs_csv(WorkItem(self.invoice, None, REASON_VALIDATED))

        self.assertTrue(path.is_relative_to(self.history_path.parent.resolve()))
        self.assertEqual(path.read_text(encoding="utf-8"), "account,amount\n1,100\n")

    def test_export_csv_missing_raises(self):
        self.csv_path.unlink()

        with self.assertRaises(FileNotFoundError):
            allocated_costs_csv(WorkItem(self.invoice, self.bundle, None))


class ExportDocumentRulesTests(unittest.TestCase):
    def test_only_fail_or_error_notifications_are_errors(self):
        notifications = [
            {"text": "Saved", "type": "success"},
            {"text": "Vat Code is not valid", "type": "fail"},
            {"text": "Server error", "type": "error"},
            {"text": "Check data", "type": "warning"},
        ]

        self.assertEqual(error_messages(notifications), ["Vat Code is not valid", "Server error"])
    def test_vat_code_depends_on_total_vat_amount(self):
        for value in ("0", "0,00", "0.00", " 0 "):
            self.assertEqual(vat_code_for(value), VAT_CODE_WITHOUT_VAT, value)
        for value in ("1950686", "1.950.686,00", "1,950,686.00", "0,01"):
            self.assertEqual(vat_code_for(value), VAT_CODE_WITH_VAT, value)

    def test_vat_code_rejects_unreadable_or_negative_amount(self):
        for value in ("", "  ", "-5"):
            with self.assertRaises(ValueError, msg=value):
                vat_code_for(value)

    def test_zero_balance_pattern(self):
        for text in ("Balance:0", "Debit 10 Balance:0 Credit 10", "Balance: 0,00", "Balance:0.00"):
            self.assertRegex(text, ZERO_BALANCE)
        for text in ("Balance:0.5", "Balance:05", "Balance:10", "Balance:0,01", "Balance:-3"):
            self.assertNotRegex(text, ZERO_BALANCE)


if __name__ == "__main__":
    unittest.main()
