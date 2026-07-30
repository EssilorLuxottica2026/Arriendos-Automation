from pathlib import Path
import unittest

from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader
from lease_accounting.pipeline.processor import LeaseAccountingPipeline


FIXTURES = Path(__file__).parent / "fixtures"


class DocumentChargeTests(unittest.TestCase):
    def test_document_interest_becomes_an_allocated_cost_item(self):
        parsed = InvoiceSupportReader().parse_file(FIXTURES / "invoice_document_interest.xml")

        interest = next(item for item in parsed.line_items if item.get("source") == "xml_document_charge")
        self.assertEqual(interest["description"], "Intereses")
        self.assertEqual(interest["amount"], 25222)
        self.assertEqual(sum(item["amount"] for item in parsed.line_items), parsed.total)

        pipeline = LeaseAccountingPipeline("scratch")
        concept, account = pipeline._item_line_classification(
            interest,
            {"account": pipeline.FIXED_ACCOUNT, "rent_type": "RF"},
        )
        self.assertEqual(concept, "INTERES")
        self.assertEqual(account, "1537210004")


if __name__ == "__main__":
    unittest.main()
