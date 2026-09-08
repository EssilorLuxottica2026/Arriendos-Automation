from pathlib import Path
import unittest

from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader


FIXTURES = Path(__file__).parent / "fixtures"


class InvoiceNoteTests(unittest.TestCase):
    def test_document_note_can_declare_a_discount(self):
        parsed = InvoiceSupportReader().parse_file(FIXTURES / "invoice_note_discount.xml")

        self.assertEqual(parsed.document_type, "Invoice")
        self.assertEqual(len(parsed.discounts), 1)
        self.assertEqual(parsed.discounts[0]["source"], "xml_note")
        self.assertEqual(parsed.discounts[0]["amount"], 100)
        self.assertEqual(parsed.discounts[0]["condition_type"], "explicit_deadline")
