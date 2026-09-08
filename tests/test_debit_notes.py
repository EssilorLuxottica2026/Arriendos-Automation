from pathlib import Path
import unittest

from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader


FIXTURES = Path(__file__).parent / "fixtures"


class DebitNoteTests(unittest.TestCase):
    def test_uses_requested_monetary_total(self):
        parsed = InvoiceSupportReader().parse_file(FIXTURES / "debit_note.xml")

        self.assertEqual(parsed.document_type, "DebitNote")
        self.assertEqual(parsed.subtotal, 900)
        self.assertEqual(parsed.total, 119)
        self.assertEqual(parsed.line_items[0]["description"], "RECARGO ADMINISTRATIVO")
        self.assertEqual(parsed.line_items[0]["quantity"], 2)
