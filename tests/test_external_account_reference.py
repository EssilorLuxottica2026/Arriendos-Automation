import tempfile
import unittest
from pathlib import Path

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline


class ExternalAccountReferenceTests(unittest.TestCase):
    def test_account_concepts_are_loaded_from_excel_reference(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "FACTURAS_CONTABILIZADAS.xlsx"
            pd.DataFrame(
                [
                    {"Cuenta": "1245150017", "Descripcion": "ARRIENDOS FIJOS", "Concepto": "RF"},
                    {"Cuenta": "1537210004", "Descripcion": "ARRIENDO VARIABLE", "Concepto": "RV"},
                    {"Cuenta": "1532001605", "Descripcion": "ENERGIA", "Concepto": "ENERGÍA"},
                    {"Cuenta": "1532001604", "Descripcion": "AGUA Y DRENAJE", "Concepto": "ACUED"},
                ]
            ).to_excel(workbook_path, index=False, sheet_name="Conceptos")

            pipeline = LeaseAccountingPipeline("scratch", account_reference_path=workbook_path)

            self.assertEqual(pipeline._resolve_account_concept("1245150017", "ARRIENDOS FIJOS"), "RF")
            self.assertEqual(pipeline._resolve_account_concept("1537210004", "ARRIENDO VARIABLE"), "RV")
            self.assertEqual(pipeline._resolve_account_concept("1532001605", "ENERGIA"), "ENERGÍA")
            self.assertEqual(pipeline._resolve_account_concept("1532001604", "AGUA Y DRENAJE"), "ACUED")

    def test_real_history_layout_infers_concepts_from_account_description(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "FACTURAS_CONTABILIZADAS.xlsx"
            pd.DataFrame(
                [
                    ["CUENTAS DE MAYOR PARA ARRIENDOS", ""],
                    ["CUENTA", "NOMBRE DE CUENTA"],
                    ["1245150017", "ARRIENDOS FIJOS (Si esta vigente)"],
                    ["1537210004", "ARRIENDO VARIABLE (Cancelado/Cesado/cualquier otro)"],
                    ["1149120011", "CUENTA DEL IVA"],
                ]
            ).to_excel(workbook_path, index=False, header=False, sheet_name="CONSOLIDADO DE CUENTAS ARRI")

            pipeline = LeaseAccountingPipeline("scratch", account_reference_path=workbook_path)

            self.assertEqual(pipeline.FIXED_ACCOUNT, "1245150017")
            self.assertEqual(pipeline.VARIABLE_ACCOUNT, "1537210004")
            self.assertEqual(pipeline.VAT_ACCOUNT, "1149120011")


if __name__ == "__main__":
    unittest.main()
