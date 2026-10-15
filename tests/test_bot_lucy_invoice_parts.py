import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from lucy.bundle import InvoiceBundle, read_bundle
from lucy.processes import invoice_parts
from lucy.processes.allocated_costs_review import AllocatedCostsReviewError, invoice_total
from lucy.processes.invoice_parts import (
    COPY_RESULT_KEY,
    PART_ERROR,
    PART_POSTED,
    PART_REVIEWED,
    PartsError,
    confirm_all_posted,
    is_part_done,
    mark_part,
    part_item,
    part_vendor_code,
    pick_copy_index,
    preflight_parts,
    single_part_view,
)

HEADER = (
    "D/A$S/H,Imp$Amount,Cod_Iva$Vat Code,Conto$GL Account,CdC$Cost Center,ProfitCenter$Profit Center,"
    "Text$Text,Brand$Brand"
)


def csv_text(base, ceco, store):
    return "\n".join(
        [
            HEADER,
            f"40,{base},VQ,1537210004,,{ceco},10-2026 INTERES {store} {ceco},{store}",
        ]
    )


def make_part(number, vendor, base, ceco, store):
    return {
        "parte": number,
        "vendor": vendor,
        "tienda": store,
        "ceco": ceco,
        "importe": base,
        "base": base,
        "iva": 0,
        "barcode": "8140230027762",
        "csv": {"allocated_costs": f"csv\\part_{number}.csv"},
        "pasos_confirmados": ["preparacion_completada"],
    }


class InvoicePartsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.directory = Path(self.tmp.name)
        (self.directory / "csv").mkdir()
        (self.directory / "csv" / "part_1.csv").write_text(csv_text(1000, "6554", "T004"), encoding="utf-8")
        (self.directory / "csv" / "part_2.csv").write_text(csv_text(2000, "6555", "T005"), encoding="utf-8")
        self.invoice = {
            "folio": "FE1",
            "vendors": ["204532", "204533"],
            "importe": 3000,
            "partes": [make_part(1, "204532", 1000, "6554", "T004"), make_part(2, "204533", 2000, "6555", "T005")],
        }
        self.json_path = self.directory / "control_facturas_bot.json"
        self.json_path.write_text(json.dumps({"facturas": [self.invoice]}), encoding="utf-8")
        self.bundle = read_bundle(self.directory)
        self.invoice = self.bundle.data["facturas"][0]

    def test_vendor_code_is_padded(self):
        self.assertEqual(part_vendor_code(self.invoice["partes"][0]), "0000204532")

    def test_single_part_view_keeps_only_that_part_and_its_amounts(self):
        part = self.invoice["partes"][1]
        view = single_part_view(self.invoice, part)
        self.assertEqual(view["partes"], [part])
        self.assertEqual(view["vendors"], ["204533"])
        self.assertEqual(invoice_total(view, "importe"), 2000)
        self.assertEqual(len(self.invoice["partes"]), 2)

    def test_part_item_resolves_the_csv_of_its_own_part(self):
        from lucy.processes.export_document import allocated_costs_csv

        path = allocated_costs_csv(part_item(self.bundle, self.invoice, self.invoice["partes"][1]))
        self.assertEqual(path.name, "part_2.csv")

    def test_pick_copy_when_all_copies_stay_in_the_list(self):
        self.assertEqual(pick_copy_index(3, 3, 0, dry_run=False), 0)
        self.assertEqual(pick_copy_index(3, 3, 1, dry_run=False), 1)
        self.assertEqual(pick_copy_index(3, 3, 2, dry_run=False), 2)

    def test_pick_copy_when_processed_copies_leave_the_list(self):
        self.assertEqual(pick_copy_index(2, 3, 1, dry_run=False), 0)
        self.assertEqual(pick_copy_index(1, 3, 2, dry_run=False), 0)

    def test_pick_copy_rejects_unexpected_counts(self):
        with self.assertRaises(PartsError):
            pick_copy_index(5, 3, 1, dry_run=False)
        with self.assertRaises(PartsError):
            pick_copy_index(1, 3, 0, dry_run=False)

    def test_pick_copy_in_dry_run_uses_the_available_copy(self):
        self.assertEqual(pick_copy_index(1, 2, 0, dry_run=True), 0)
        self.assertEqual(pick_copy_index(1, 2, 1, dry_run=True), 0)

    def test_preflight_accepts_valid_parts_in_dry_run(self):
        with mock.patch.object(invoice_parts, "REPOST_DRY_RUN", True):
            preflight_parts(self.bundle, self.invoice)

    def test_preflight_blocks_real_mode_until_sap_posting_exists(self):
        with mock.patch.object(invoice_parts, "REPOST_DRY_RUN", False):
            with self.assertRaises(PartsError):
                preflight_parts(self.bundle, self.invoice)

    def test_preflight_rejects_parts_that_do_not_match_vendors(self):
        self.invoice["vendors"] = ["204532", "204533", "204534"]
        with mock.patch.object(invoice_parts, "REPOST_DRY_RUN", True):
            with self.assertRaises(PartsError):
                preflight_parts(self.bundle, self.invoice)

    def test_preflight_rejects_a_csv_that_does_not_match_its_part(self):
        self.invoice["partes"][0]["ceco"] = "9999"
        with mock.patch.object(invoice_parts, "REPOST_DRY_RUN", True):
            with self.assertRaises(AllocatedCostsReviewError):
                preflight_parts(self.bundle, self.invoice)

    def test_mark_part_persists_result_in_the_json(self):
        part = self.invoice["partes"][0]
        mark_part(self.bundle, part, PART_REVIEWED, {"resumen": "x"})
        saved = json.loads(self.json_path.read_text(encoding="utf-8"))["facturas"][0]["partes"][0]
        self.assertEqual(saved[COPY_RESULT_KEY]["estado"], PART_REVIEWED)
        self.assertEqual(saved[COPY_RESULT_KEY]["resumen_allocated_cost"], {"resumen": "x"})
        self.assertIn("csv_importado_y_revisado", saved["pasos_confirmados"])
        self.assertNotIn("contabilizada_sap", saved["pasos_confirmados"])
        self.assertTrue(is_part_done(part))

    def test_error_part_is_not_done_and_keeps_the_error(self):
        part = self.invoice["partes"][0]
        mark_part(self.bundle, part, PART_ERROR, error="boom")
        self.assertFalse(is_part_done(part))
        self.assertEqual(part[COPY_RESULT_KEY]["error"], "boom")
        self.assertEqual(part["pasos_confirmados"], ["preparacion_completada"])

    def test_confirm_all_posted(self):
        with mock.patch.object(invoice_parts, "REPOST_DRY_RUN", False):
            with self.assertRaises(PartsError):
                confirm_all_posted(self.invoice)
            for part in self.invoice["partes"]:
                mark_part(self.bundle, part, PART_POSTED)
            confirm_all_posted(self.invoice)
        self.assertIn("contabilizada_sap", self.invoice["partes"][0]["pasos_confirmados"])

    def test_confirm_all_posted_in_dry_run_does_not_fail(self):
        with mock.patch.object(invoice_parts, "REPOST_DRY_RUN", True):
            confirm_all_posted(self.invoice)


if __name__ == "__main__":
    unittest.main()
