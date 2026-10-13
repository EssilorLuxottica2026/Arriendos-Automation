from decimal import Decimal
from pathlib import Path
import tempfile
import unittest

from lucy.processes.allocated_costs_review import (
    AllocatedCostsReviewError,
    accounting_issues,
    check_before_import,
    concept_summary,
    deductible_vat,
    grid_issues,
    line_concept,
    parse_amount,
    summary_issues,
)

HEADER = (
    "D/A$S/H,Imp$Amount,Cod_Iva$Vat Code,Conto$GL Account,CdC$Cost Center,ProfitCenter$Profit Center,"
    "Text$Text,Ordine$Internal Order,Wbs$WBS,Attribuzione$Assignment,Network$Network,OpNetwork$OpNetwork,"
    "MaterialNumber$Material Number,Quantity$Quantity,TransactionType$Transaction Type,POR$POR#,"
    "CompanyGL$CompanyGL,Brand$Brand,TradingPartner$Trading Partner,PosOrd$Order Position,OdaOdv$PO/SR"
)


def row(sh, amount, vat, account, text, ceco="6554", store="T004"):
    return {
        "D/A$S/H": sh,
        "Imp$Amount": str(amount),
        "Cod_Iva$Vat Code": vat,
        "Conto$GL Account": account,
        "CdC$Cost Center": "",
        "ProfitCenter$Profit Center": ceco,
        "Text$Text": text,
        "Brand$Brand": store,
    }


PART = {"vendor": "204532", "tienda": "T004", "ceco": "6554", "importe": 58285, "base": 48979, "iva": 9306}
ROWS = [
    row("40", 48979, "VQ", "1537210004", "10-2026 INTERES T004 6554"),
    row("40", 354, "VW", "1149120011", "10-2026 RV T004 6554"),
    row("50", 354, "VQ", "1149120011", "10-2026 RV T004 6554"),
]


class AllocatedCostsReviewTests(unittest.TestCase):
    def test_parse_amount_formats(self):
        for text, expected in (
            ("48979", "48979"),
            ("48.979", "48979"),
            ("48.979,00", "48979.00"),
            ("$ 48,979.50", "48979.50"),
            ("1.234.567", "1234567"),
            ("12,5", "12.5"),
        ):
            self.assertEqual(parse_amount(text), Decimal(expected), text)
        with self.assertRaises(ValueError):
            parse_amount("")

    def test_concept_keeps_multiword_sigla(self):
        self.assertEqual(line_concept(row("40", 1, "VW", "1537210002", "10-2026 GC V SGH608 24", "24", "SGH608")), "GC V")
        self.assertEqual(line_concept(ROWS[0]), "INTERES")

    def test_concept_with_multiword_store(self):
        office = [row("40", 1, "VQ", "1245150017", "09-2026 RF OFICINA T3 8140074", "8140074", "OFICINA T3")]
        self.assertEqual(line_concept(office[0]), "RF")
        self.assertEqual(concept_summary(office), "09-2026 RF OFICINA T3 8140074")
        self.assertEqual(summary_issues("09-2026 RF OFICINA T3 8140074", office), [])

    def test_summary_lists_unique_concepts_with_slash(self):
        self.assertEqual(concept_summary(ROWS), "10-2026 INTERES/RV T004 6554")

    def test_summary_check_reports_missing_and_repeated(self):
        self.assertEqual(summary_issues("10-2026 INTERES/RV T004 6554", ROWS), [])
        self.assertEqual(summary_issues("10-2026 RV/INTERES T004 6554", ROWS), [])
        issues = summary_issues("10-2026 RV/RV T004 6554", ROWS)
        self.assertIn("al resumen le faltan conceptos: INTERES", issues)
        self.assertIn("el resumen repite conceptos: RV", issues)

    def test_summary_requires_single_store_and_ceco(self):
        with self.assertRaises(AllocatedCostsReviewError):
            concept_summary([ROWS[0], row("40", 1, "VQ", "1537210004", "10-2026 RF T005 6555", "6555", "T005")])

    def test_accounting_ok_with_deductible_vat_pair(self):
        self.assertEqual(accounting_issues(ROWS, PART), [])
        self.assertEqual(deductible_vat(ROWS), Decimal(354))

    def test_accounting_ok_without_deductible_vat(self):
        rows = [
            row("40", 16890000, "VW", "1537210002", "10-2026 GC V SGH608 24", "24", "SGH608"),
            row("40", 2146409, "VW", "1532001600", "10-2026 ASEO SGH608 24", "24", "SGH608"),
        ]
        part = {"tienda": "SGH608", "ceco": "24", "base": 19036409, "iva": 3616918}
        self.assertEqual(accounting_issues(rows, part), [])
        self.assertEqual(deductible_vat(rows), 0)

    def test_accounting_detects_wrong_ceco_account_and_base(self):
        bad = [dict(ROWS[0], **{"ProfitCenter$Profit Center": "9999", "Conto$GL Account": "15372"})] + ROWS[1:]
        issues = accounting_issues(bad, PART)
        self.assertTrue(any("CeCo '9999'" in issue for issue in issues))
        self.assertTrue(any("cuenta '15372'" in issue for issue in issues))

        issues = accounting_issues([dict(ROWS[0], **{"Imp$Amount": "48000"})] + ROWS[1:], PART)
        self.assertTrue(any("neto de las líneas" in issue for issue in issues))

    def test_accounting_detects_unmatched_vat_reversal(self):
        rows = [ROWS[0], ROWS[1], dict(ROWS[2], **{"Imp$Amount": "300"}), row("40", 54, "VQ", "1537210004", "10-2026 INTERES T004 6554")]
        issues = accounting_issues(rows, PART)
        self.assertTrue(any("no compensa ningún IVA deducible" in issue for issue in issues))

    def test_grid_matches_lucy_formats_in_any_order(self):
        grid = [
            ["2", "50", "354,00", "VQ - IVA Compra Servicios 19% - Gasto", "1149120011 - IVA", "", "6554 - Tienda", "10-2026 RV T004 6554"],
            ["1", "40", "48.979,00", "VQ", "1537210004", "", "6554", "10-2026 INTERES T004 6554"],
            ["3", "40", "354", "VW", "1149120011", "", "6554", "10-2026 RV T004 6554"],
        ]
        self.assertEqual(grid_issues(ROWS, grid), [])

    def test_grid_detects_changed_text_and_missing_line(self):
        grid = [
            ["1", "40", "48.979,00", "VQ", "1537210004", "", "6554", "10-2026 INTERES T004"],
            ["3", "40", "354", "VW", "1149120011", "", "6554", "10-2026 RV T004 6554"],
        ]
        issues = grid_issues(ROWS, grid)
        self.assertIn("Lucy muestra 2 líneas y el CSV tiene 3", issues)
        self.assertEqual(sum("no aparece igual en Lucy" in issue for issue in issues), 2)

    def test_check_before_import_reads_real_csv_format(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "F1_allocated_costs.csv"
            lines = [HEADER] + [
                f"{r['D/A$S/H']},{r['Imp$Amount']},{r['Cod_Iva$Vat Code']},{r['Conto$GL Account']},,"
                f"{r['ProfitCenter$Profit Center']},{r['Text$Text']},,,F1,,,,,,,,{r['Brand$Brand']},,,"
                for r in ROWS
            ]
            path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
            rows = check_before_import("F1", PART, path)
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[0]["Text$Text"], "10-2026 INTERES T004 6554")

            with self.assertRaises(AllocatedCostsReviewError):
                check_before_import("F1", dict(PART, base=1), path)


if __name__ == "__main__":
    unittest.main()
