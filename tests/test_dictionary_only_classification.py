import unittest

import pandas as pd

from lease_accounting.pipeline.processor import LeaseAccountingPipeline, PipelineError


class DictionaryOnlyClassificationTests(unittest.TestCase):
    def setUp(self):
        self.pipeline = LeaseAccountingPipeline("scratch")

    def test_manto_is_not_assigned_by_hardcoded_keyword(self):
        item = {"description": "MANTO ZONA COMUN", "amount": 500}

        concept, account = self.pipeline._item_line_classification(
            item,
            {"account": "1245150017", "rent_type": "RF"},
        )

        self.assertEqual(concept, "SIN CLASIFICAR")
        self.assertEqual(account, "")

    def test_rent_keyword_also_requires_dictionary_learning(self):
        output_df = pd.DataFrame(
            [
                {
                    "invoice_id": "R1",
                    "store": "T001",
                    "ceco": "6001",
                    "account": "1245150017",
                    "rent_type": "RF",
                    "invoice_line_items": [
                        {"description": "CANON VARIABLE", "amount": 1000}
                    ],
                }
            ]
        )

        review_items = self.pipeline.build_concept_review_items(output_df)

        self.assertEqual(len(review_items), 1)
        self.assertEqual(review_items[0]["description"], "CANON VARIABLE")

    def test_learned_phrase_uses_user_concept_and_account(self):
        self.pipeline.learning_dictionary = pd.DataFrame(
            [
                {
                    "concept": "OTRO",
                    "account": "9999999999",
                    "phrases": "MANTO ZONA COMUN",
                    "text": "OTRO",
                }
            ]
        )

        concept, account = self.pipeline._item_line_classification(
            {"description": "MANTO ZONA COMUN", "amount": 500},
            {"account": "1245150017", "rent_type": "RF"},
        )

        self.assertEqual(concept, "OTRO")
        self.assertEqual(account, "9999999999")

    def test_generation_is_blocked_while_items_are_unlearned(self):
        output_df = pd.DataFrame(
            [
                {
                    "invoice_id": "M1",
                    "store": "T001",
                    "ceco": "6001",
                    "invoice_line_items": [
                        {"description": "MANTO ZONA COMUN", "amount": 500}
                    ],
                }
            ]
        )

        with self.assertRaisesRegex(PipelineError, "sin concepto aprendido"):
            self.pipeline.generate_output(
                output_df,
                pd.DataFrame(),
                pd.DataFrame(),
            )


if __name__ == "__main__":
    unittest.main()
