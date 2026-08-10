from pathlib import Path
import unittest

from lease_accounting.pipeline.pdf_reader import InvoiceSupportReader
from lease_accounting.pipeline.processor import LeaseAccountingPipeline


FIXTURES = Path(__file__).parent / "fixtures"


class DocumentChargeTests(unittest.TestCase):
    def test_document_interest_requires_dictionary_learning(self):
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
        self.assertEqual(concept, "SIN CLASIFICAR")
        self.assertEqual(account, "")

        review_items = pipeline.build_concept_review_items(
            __import__("pandas").DataFrame(
                [
                    {
                        "invoice_id": parsed.invoice_id,
                        "store": "T078",
                        "ceco": "6700",
                        "account": pipeline.FIXED_ACCOUNT,
                        "rent_type": "RF",
                        "invoice_line_items": [interest],
                    }
                ]
            )
        )
        self.assertEqual(len(review_items), 1)
        self.assertEqual(review_items[0]["description"], "Intereses")


if __name__ == "__main__":
    unittest.main()
