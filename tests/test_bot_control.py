import csv
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline, PipelineError
from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader
from lease_accounting.web import app


FIXTURES = Path(__file__).parent / "fixtures"


class BotControlTests(unittest.TestCase):
    def setUp(self):
        root = Path.cwd() / "scratch"
        root.mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=root)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.pipeline = LeaseAccountingPipeline(self.root / "output")
        self.support = self.root / "original  factura.xml"
        self.write_xml(self.support, "F1", 100)

    def write_xml(self, path, folio, gross, vat=0):
        xml = (FIXTURES / "credit_note.xml").read_text(encoding="utf-8")
        xml = xml.replace("CreditNote", "Invoice").replace("CreditedQuantity", "InvoicedQuantity")
        xml = xml.replace("FENC17661", folio).replace("1000.00", str(gross))
        xml = xml.replace("190.00", str(vat)).replace("1190.00", str(gross + vat))
        xml = xml.replace("2026-07-15", "2026-10-01")
        path.write_text(xml, encoding="utf-8")

    def posting(self, invoice_id="F1", vendor="205766", amount=100, **changes):
        return {
            "invoice_id": invoice_id,
            "vendor": vendor,
            "store": "T001",
            "ceco": "6001",
            "profit_center": "6001",
            "account": "1245150017",
            "concept": "GC",
            "rent_type": "RF",
            "amount": amount,
            "gross_amount": amount,
            "discount_total": 0,
            "posting_discount_total": 0,
            "payable_rounding": 0,
            "invoice_total": amount,
            "vat_total": 0,
            "vat_vw": 0,
            "vat_vq": 0,
            "vw_percent": 0,
            "vq_percent": 1,
            "withholding_tax_amount": 0,
            "text": "10-2026 GC T001 6001",
            "invoice_date": pd.Timestamp("2026-10-01"),
            "due_date": pd.NaT,
            "payment_terms_text": None,
            "item_count_hint": 0,
            "ubl_document_type": "Invoice",
            "invoice_line_items": [],
            "invoice_discounts": [],
            "support_source_path": str(self.support.resolve()),
            "contract_status_blocked": False,
            **changes,
        }

    def generate(self, rows, validation=None, align_xml=True):
        if align_xml:
            default_rows = [
                row for row in rows
                if row.get("support_source_path") == str(self.support.resolve())
                and row["ubl_document_type"] == "Invoice"
            ]
            for folio in dict.fromkeys(row["invoice_id"] for row in default_rows):
                group = [row for row in default_rows if row["invoice_id"] == folio]
                path = self.support if folio == "F1" else self.root / f"support_{folio.replace('/', '_').replace(':', '_')}.xml"
                self.write_xml(
                    path, folio, sum(row["gross_amount"] for row in group),
                    sum(row["vat_total"] for row in group),
                )
                for row in group:
                    row["support_source_path"] = str(path.resolve())
        result = self.pipeline.generate_output(
            pd.DataFrame(rows),
            validation if validation is not None else pd.DataFrame(),
            pd.DataFrame(),
        ).to_dict()
        path = self.pipeline.base_output_dir.parent / result["bot_control_json"]
        return result, json.loads(path.read_text(encoding="utf-8"))

    def test_links_each_part_to_exact_csv_rows_and_copied_support(self):
        result, payload = self.generate([
            self.posting(amount=125, lucy_id="00012", barcode="00099"),
            self.posting(amount=75, store="T002", ceco="6002", profit_center="6002"),
            self.posting(invoice_id="F2", amount=250),
        ])
        self.assertEqual(payload["schema_version"], 2)
        self.assertEqual(payload["rutas_habilitadas"], ["single_vendor"])
        self.assertEqual(len(payload["facturas"]), 2)
        invoice = payload["facturas"][0]
        self.assertEqual(invoice["importe"], 200)
        self.assertEqual(invoice["vendor"], "205766")
        self.assertEqual(invoice["ruta"], "single_vendor")
        self.assertTrue(invoice["lista_para_bot"])
        self.assertEqual(invoice["pasos_confirmados"], [
            "preparacion_completada", "validacion_preparacion_completada",
        ])
        self.assertEqual(invoice["validaciones"]["duplicados"], "no_ejecutada")
        self.assertIsNone(invoice["documento_sap"])
        self.assertEqual(len(invoice["partes"]), 2)
        self.assertEqual(invoice["partes"][0]["identificador_lucy"], "00012")
        self.assertEqual(invoice["partes"][0]["barcode"], "00099")
        self.assertIsNone(invoice["partes"][1]["identificador_lucy"])
        self.assertIsNone(invoice["partes"][1]["barcode"])
        self.assertEqual(invoice["partes"][1]["ceco"], "6002")
        self.assertEqual(invoice["partes"][1]["importe"], 75)
        self.assertEqual(invoice["partes"][1]["moneda"], "COP")
        self.assertEqual(invoice["partes"][1]["profit_center"], "6002")

        with zipfile.ZipFile(self.pipeline.base_output_dir.parent / result["bot_control_bundle"]) as bundle:
            expected_members = {
                "control_facturas_bot.json",
                *(
                    Path(part["csv"]["allocated_costs"]).as_posix()
                    for entry in payload["facturas"]
                    for part in entry["partes"]
                ),
            }
            self.assertEqual(set(bundle.namelist()), expected_members)
            self.assertEqual(len(bundle.namelist()), len(expected_members))
            self.assertEqual(
                json.loads(bundle.read("control_facturas_bot.json").decode("utf-8")),
                payload,
            )
            for entry in payload["facturas"]:
                for part in entry["partes"]:
                    self.assertEqual(
                        (self.pipeline.output_dir / part["archivo_factura"]).read_bytes(),
                        Path(part["archivo_origen"]).read_bytes(),
                    )
                    self.assertNotIn(Path(part["archivo_factura"]).as_posix(), bundle.namelist())
                    for kind, relative_path in part["csv"].items():
                        self.assertTrue((self.pipeline.output_dir / relative_path).is_file())
                        if kind == "allocated_costs":
                            self.assertIn(Path(relative_path).as_posix(), bundle.namelist())
                            self.assertEqual(
                                bundle.read(Path(relative_path).as_posix()),
                                (self.pipeline.output_dir / relative_path).read_bytes(),
                            )
                        else:
                            self.assertNotIn(Path(relative_path).as_posix(), bundle.namelist())
                    with (self.pipeline.output_dir / part["csv"]["allocated_costs"]).open(
                        encoding="utf-8-sig", newline=""
                    ) as file:
                        rows = list(csv.DictReader(file))
                    self.assertEqual(len(rows), len(part["allocated_costs"]))
                    for exported, saved in zip(rows, part["allocated_costs"]):
                        self.assertEqual(set(exported), set(saved))
                        for column, value in saved.items():
                            self.assertEqual(exported[column], "" if value is None else str(value))
                        self.assertEqual(exported["Attribuzione$Assignment"], entry["folio"])
                        self.assertEqual(exported["CdC$Cost Center"], "")
                    self.assertNotIn("_posting_index", part["allocated_costs"][0])
                    self.assertIsNone(part["documento_sap"])
                    self.assertIsNone(part["copia_lucy"])
                    self.assertIsNone(part["compensacion_sap"])
            self.assertEqual(len(list((self.pipeline.output_dir / "soportes_facturas").iterdir())), 2)

    def test_multivendor_and_missing_support_are_not_ready(self):
        result, payload = self.generate([
            self.posting(vendor="111"),
            self.posting(vendor="222"),
            self.posting(invoice_id="F2", support_source_path=None),
        ])
        multi, missing = payload["facturas"]
        self.assertEqual(multi["vendors"], ["111", "222"])
        self.assertIsNone(multi["ruta"])
        self.assertEqual(multi["estado"], "pendiente_ruta")
        self.assertFalse(multi["lista_para_bot"])
        self.assertIn("multivendor", multi["error"])
        self.assertEqual(missing["ruta"], "single_vendor")
        self.assertEqual(missing["estado"], "pendiente_archivo")
        self.assertFalse(missing["lista_para_bot"])
        self.assertIsNone(missing["partes"][0]["archivo_factura"])
        self.assertEqual(len(result["warnings"]), 2)
        self.assertEqual(result["output_rows"], 3)

    def test_validation_errors_and_blocked_contracts_are_excluded(self):
        validation = pd.DataFrame([{
            "invoice_id": "INVALID",
            "missing_ceco": True,
            "missing_contract": False,
            "missing_prorrateo": False,
        }])
        _, payload = self.generate([
            self.posting(),
            self.posting(invoice_id="INVALID"),
            self.posting(invoice_id="BLOCKED", contract_status_blocked=True),
        ], validation)
        self.assertEqual([item["folio"] for item in payload["facturas"]], ["F1"])
        self.assertEqual(payload["incidencias_procesamiento"][0]["invoice_id"], "INVALID")

    def test_missing_source_file_fails_explicitly(self):
        with self.assertLogs("lease_accounting.pipeline.processor", level="ERROR"):
            with self.assertRaisesRegex(PipelineError, "missing.xml"):
                self.generate([self.posting(support_source_path=str(self.root / "missing.xml"))])

    def test_all_blocked_contracts_do_not_generate_an_empty_control(self):
        with self.assertRaisesRegex(PipelineError, "No hay facturas habilitadas"):
            self.generate([self.posting(contract_status_blocked=True)])
        self.assertFalse((self.pipeline.output_dir / "control_facturas_bot.json").exists())

    def test_prevents_overwriting_bot_progress_and_linked_csvs(self):
        result, _ = self.generate([self.posting()])
        path = self.pipeline.base_output_dir.parent / result["bot_control_json"]
        payload = json.loads(path.read_text(encoding="utf-8"))
        payload["facturas"][0]["documento_sap"] = "SAP-CONFIRMADO"
        path.write_text(json.dumps(payload), encoding="utf-8")
        before = path.read_bytes()
        with self.assertRaisesRegex(PipelineError, "conservar su avance"):
            self.generate([self.posting(amount=999)])
        self.assertEqual(path.read_bytes(), before)

    def test_rejects_colliding_invoice_filenames(self):
        with self.assertRaisesRegex(PipelineError, "mismo nombre"):
            self.generate([self.posting(invoice_id="F/1"), self.posting(invoice_id="F:1")])

    def test_input_identifiers_preserve_leading_zeros(self):
        path = self.root / "invoices.csv"
        pd.DataFrame([{
            "Factura": "F1", "Vendor": "00123",
            "ID_Lucy": "00012", "Codigo_de_barras": "00099",
        }]).to_csv(path, index=False)
        loaded = self.pipeline._load_invoices(path)
        self.assertEqual(loaded.iloc[0]["lucy_id"], "00012")
        self.assertEqual(loaded.iloc[0]["barcode"], "00099")
        self.assertEqual(loaded.iloc[0]["vendor"], "00123")

    def test_support_reader_retains_absolute_source_path(self):
        supports = self.pipeline._load_invoice_supports([self.support])
        self.assertEqual(supports.iloc[0]["source_path"], str(self.support.resolve()))
        direct = self.pipeline._load_invoices(self.support)
        self.assertEqual(direct.iloc[0]["invoice_source_path"], str(self.support.resolve()))

    def test_container_barcode_survives_matching_splitting_and_export(self):
        xml = (FIXTURES / "attached_credit_note.xml").read_text(encoding="utf-8")
        xml = xml.replace(
            "<cbc:ID>ATTACHED-FENC17661</cbc:ID>",
            "<cbc:ID>CONTAINER-ID-NOT-BARCODE</cbc:ID>"
            "<cbc:ParentDocumentID>004092</cbc:ParentDocumentID>",
        )
        self.support.write_text(xml, encoding="utf-8")
        parsed = InvoiceSupportReader().parse_file(self.support)
        self.assertEqual(parsed.barcode, "004092")
        self.assertEqual(parsed.invoice_id, "FENC17661")
        invoices = self.pipeline._load_invoices(self.support)
        invoices["barcode"] = "TABLE-VALUE"
        invoices["lucy_id"] = "LUCY-12"
        supports = self.pipeline._load_invoice_supports(
            [self.support], support_split_factors={str(self.support): 2}
        )
        data = {
            "invoices": invoices,
            "invoice_supports": supports,
            "control": pd.DataFrame(columns=["vendor_code", "vendor", "store", "ceco"]),
            "contracts": pd.DataFrame(columns=["ceco", "end_of_term", "rent_min", "sell_media"]),
            "prorateo": pd.DataFrame(columns=["store", "ceco", "vw_percent"]),
            "distribution": pd.DataFrame(columns=["source_vendor", "target_vendor", "split_percent"]),
            "macro_database": pd.DataFrame(),
        }
        cleaned = self.pipeline.clean_data(data)["invoices"]
        self.assertEqual(cleaned["barcode"].tolist(), ["004092", "004092"])
        self.assertEqual(cleaned["lucy_id"].tolist(), ["LUCY-12", "LUCY-12"])
        self.assertEqual(
            cleaned["support_source_path"].tolist(),
            [str(self.support.resolve())] * 2,
        )
        _, payload = self.generate([
            self.posting(
                invoice_id=row["invoice_id"], amount=250,
                barcode=row["barcode"], lucy_id=row["lucy_id"],
                support_source_path=row["support_source_path"],
                ubl_document_type="CreditNote",
                invoice_date=pd.Timestamp("2026-07-15"),
            )
            for _, row in cleaned.iterrows()
        ])
        self.assertEqual(
            [part["barcode"] for part in payload["facturas"][0]["partes"]],
            ["004092", "004092"],
        )
        self.assertEqual(payload["facturas"][0]["partes"][0]["tipo_documento"], "Supplier Credit Note")

    def test_missing_parent_document_id_is_not_replaced_by_other_ids(self):
        reader = InvoiceSupportReader()
        self.assertIsNone(reader.parse_file(FIXTURES / "attached_credit_note.xml").barcode)
        self.assertIsNone(reader.parse_file(FIXTURES / "credit_note.xml").barcode)

    def test_standalone_xml_parent_document_id_is_read(self):
        xml = (FIXTURES / "credit_note.xml").read_text(encoding="utf-8")
        xml = xml.replace(
            "<cbc:ID>FENC17661</cbc:ID>",
            "<cbc:ID>FENC17661</cbc:ID><cbc:ParentDocumentID>4092</cbc:ParentDocumentID>",
            1,
        )
        self.support.write_text(xml, encoding="utf-8")
        self.assertEqual(InvoiceSupportReader().parse_file(self.support).barcode, "4092")
        self.assertEqual(self.pipeline._load_invoices(self.support).iloc[0]["barcode"], "4092")

    def test_successive_runs_in_same_minute_keep_separate_controls(self):
        with patch.object(
            self.pipeline, "prepare",
            return_value=({}, pd.DataFrame([self.posting()]), pd.DataFrame()),
        ):
            first = self.pipeline.run(*([self.support] * 5))
            second = self.pipeline.run(*([self.support] * 5))
        self.assertNotEqual(first["bot_control_json"], second["bot_control_json"])
        for result in [first, second]:
            self.assertTrue((self.pipeline.base_output_dir / result["bot_control_json"]).is_file())

    def test_results_page_exposes_bot_downloads(self):
        result, _ = self.generate([self.posting()])
        with app.test_request_context():
            html = app.jinja_env.get_template("result.html").render(result=result)
        self.assertIn(result["bot_control_json"], html)
        self.assertIn(result["bot_control_bundle"], html)
        with patch("lease_accounting.web.OUTPUT_DIR", self.pipeline.base_output_dir.parent):
            response = app.test_client().get(f"/download/{result['bot_control_json']}")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(len(response.get_json()["facturas"]), 1)
            response.close()

    def test_xml_with_other_folio_is_pending_review(self):
        self.write_xml(self.support, "OTHER", 100)
        result, payload = self.generate([self.posting()], align_xml=False)
        invoice = payload["facturas"][0]
        self.assertFalse(invoice["lista_para_bot"])
        self.assertEqual(invoice["estado"], "pendiente_revision")
        self.assertIn("folio del XML", invoice["error"])
        self.assertEqual(invoice["partes"][0]["estado"], "pendiente_revision")
        self.assertEqual(invoice["pasos_confirmados"], ["preparacion_completada"])
        self.assertTrue(result["warnings"])

    def test_invalid_xml_is_saved_pending_review(self):
        self.support.write_text("<Invoice>", encoding="utf-8")
        _, payload = self.generate([self.posting()], align_xml=False)
        invoice = payload["facturas"][0]
        self.assertFalse(invoice["lista_para_bot"])
        self.assertIn("no se pudo leer el XML", invoice["error"])

    def test_pdf_is_not_copied_or_enabled(self):
        pdf = self.root / "unsupported.pdf"
        pdf.write_bytes(b"not a supported invoice")
        result, payload = self.generate([self.posting(support_source_path=str(pdf.resolve()))])
        invoice = payload["facturas"][0]
        self.assertFalse(invoice["lista_para_bot"])
        self.assertIn("debe ser XML", invoice["error"])
        self.assertIsNone(invoice["partes"][0]["archivo_factura"])
        with zipfile.ZipFile(self.pipeline.base_output_dir.parent / result["bot_control_bundle"]) as bundle:
            self.assertFalse(any(name.lower().endswith(".pdf") for name in bundle.namelist()))

    def test_missing_or_invalid_required_data_prevents_readiness(self):
        for field, value, message in [
            ("store", None, "tienda"),
            ("ceco", "VARIABLE", "CeCo invalido"),
            ("vendor", "", "vendor"),
            ("amount", 0, "mayor que cero"),
            ("invoice_date", pd.NaT, "fecha de factura"),
            ("account", "", "cuenta contable"),
        ]:
            with self.subTest(field=field):
                self.pipeline = LeaseAccountingPipeline(self.root / f"case_{field}")
                row = self.posting(**{field: value})
                if field == "amount":
                    row["gross_amount"] = 0
                _, payload = self.generate([row])
                invoice = payload["facturas"][0]
                self.assertFalse(invoice["lista_para_bot"])
                self.assertIn(message, invoice["error"])

    def test_xml_amount_and_tax_must_match_sum_of_parts(self):
        self.write_xml(self.support, "F1", 100, vat=19)
        _, payload = self.generate([
            self.posting(amount=60),
            self.posting(amount=50),
        ], align_xml=False)
        invoice = payload["facturas"][0]
        self.assertFalse(invoice["lista_para_bot"])
        self.assertIn("valor bruto", invoice["error"])
        self.assertIn("IVA", invoice["error"])

    def test_xml_barcode_and_type_must_match(self):
        xml = self.support.read_text(encoding="utf-8").replace(
            "<cbc:ID>F1</cbc:ID>",
            "<cbc:ID>F1</cbc:ID><cbc:ParentDocumentID>004092</cbc:ParentDocumentID>",
            1,
        )
        self.support.write_text(xml, encoding="utf-8")
        _, payload = self.generate([
            self.posting(barcode="WRONG", ubl_document_type="CreditNote")
        ], align_xml=False)
        invoice = payload["facturas"][0]
        self.assertFalse(invoice["lista_para_bot"])
        self.assertIn("barcode", invoice["error"])
        self.assertIn("tipo de documento", invoice["error"])

    def test_xml_date_mismatch_requires_review(self):
        _, payload = self.generate([
            self.posting(invoice_date=pd.Timestamp("2026-10-02"))
        ])
        self.assertIn("fecha de factura no coincide", payload["facturas"][0]["error"])

    def test_ambiguous_xml_relation_requires_review(self):
        second = self.root / "second.xml"
        self.write_xml(second, "F1", 100)
        _, payload = self.generate([
            self.posting(),
            self.posting(support_source_path=str(second.resolve())),
        ])
        self.assertFalse(payload["facturas"][0]["lista_para_bot"])
        self.assertIn("varios XML", payload["facturas"][0]["error"])

    def test_modified_allocated_costs_and_header_are_detected(self):
        for method_name, pattern, column, value, message in [
            ("_write_allocated_costs_csv_bundle", "allocated_costs_csvs_*",
             "Attribuzione$Assignment", "OTHER", "otro folio"),
            ("_write_allocated_costs_csv_bundle", "allocated_costs_csvs_*",
             "ProfitCenter$Profit Center", "OTHER", "profit center"),
            ("_write_allocated_costs_csv_bundle", "allocated_costs_csvs_*",
             "Imp$Amount", "999", "no coincide"),
            ("_write_invoice_csv_bundle", "invoice_csvs_*",
             "Invoice Total Amount", "999", "no corresponde"),
        ]:
            with self.subTest(column=column):
                self.pipeline = LeaseAccountingPipeline(self.root / f"csv_{column.replace('$', '_')}")
                original = getattr(self.pipeline, method_name)

                def tamper(**kwargs):
                    original(**kwargs)
                    directory = next(self.pipeline.output_dir.glob(pattern))
                    path = next(directory.rglob("*_header.csv" if column == "Invoice Total Amount" else "*.csv"))
                    with path.open(encoding="utf-8-sig", newline="") as file:
                        reader = csv.DictReader(file)
                        fieldnames = reader.fieldnames
                        rows = list(reader)
                    rows[0][column] = value
                    with path.open("w", encoding="utf-8-sig", newline="") as file:
                        writer = csv.DictWriter(file, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(rows)

                with patch.object(self.pipeline, method_name, side_effect=tamper):
                    result, payload = self.generate([self.posting()])
                invoice = payload["facturas"][0]
                self.assertFalse(invoice["lista_para_bot"])
                self.assertIn(message, invoice["error"])
                self.assertTrue(result["warnings"])

    def test_modified_manual_style_csv_is_detected(self):
        original = self.pipeline._write_manual_style_csv_bundle

        def tamper(**kwargs):
            original(**kwargs)
            path = next(self.pipeline.output_dir.glob("invoice_manual_style_csvs_*")).joinpath("F1_manual_style.csv")
            path.write_text(path.read_text(encoding="utf-8-sig").replace("F1", "OTHER"), encoding="utf-8-sig")

        with patch.object(self.pipeline, "_write_manual_style_csv_bundle", side_effect=tamper):
            _, payload = self.generate([self.posting()])
        self.assertFalse(payload["facturas"][0]["lista_para_bot"])
        self.assertIn("CSV de factura contiene otro folio", payload["facturas"][0]["error"])

    def test_preparation_validation_does_not_check_prior_runs_or_catalog(self):
        previous = self.pipeline.base_output_dir / "old"
        previous.mkdir()
        (previous / "control_facturas_bot.json").write_text("invalid previous JSON", encoding="utf-8")
        self.pipeline.account_reference = pd.DataFrame([{"folio": "F1", "documento_sap": "SAP"}])
        _, payload = self.generate([self.posting()])
        self.assertTrue(payload["facturas"][0]["lista_para_bot"])
        self.assertEqual(payload["verificacion_duplicados"], "no_ejecutada")
        self.assertEqual(payload["resumen_validacion"], {
            "total_facturas": 1, "preparacion_validada": 1, "pendientes": 0,
        })

    def test_valid_discount_does_not_trigger_xml_amount_mismatch(self):
        _, payload = self.generate([
            self.posting(
                amount=90, gross_amount=100, discount_total=10,
                posting_discount_total=10, vat_total=19, vat_vq=19,
            ),
        ])
        invoice = payload["facturas"][0]
        self.assertTrue(invoice["lista_para_bot"], invoice["error"])
        self.assertEqual(invoice["importe"], 109)

    def test_missing_csv_fails_instead_of_publishing_success(self):
        original = self.pipeline._write_allocated_costs_csv_bundle

        def remove_csv(**kwargs):
            original(**kwargs)
            path = next(self.pipeline.output_dir.glob("allocated_costs_csvs_*")).joinpath("F1_allocated_costs.csv")
            path.unlink()

        with (
            patch.object(self.pipeline, "_write_allocated_costs_csv_bundle", side_effect=remove_csv),
            self.assertLogs("lease_accounting.pipeline.processor", level="ERROR"),
            self.assertRaisesRegex(PipelineError, "No existe el CSV vinculado"),
        ):
            self.generate([self.posting()])
        self.assertFalse((self.pipeline.output_dir / "control_facturas_bot.json").exists())


if __name__ == "__main__":
    unittest.main()
