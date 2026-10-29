"""Unit tests for the Loughran-McDonald Sentiment Analyzer Module.

Tests dictionary loading, text cleaning, sentiment metrics calculation,
Excel 'Sentiment Analysis' tab generation, and verification.
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl

from sentiment_analyzer import (
    clean_mda_text,
    create_sample_mda_files,
    load_loughran_mcdonald_dictionary,
    process_mda_directory,
    run_sentiment_pipeline,
    score_mda_sentiment,
    verify_sentiment_tab,
    write_sentiment_to_excel,
)


class TestSentimentAnalyzer(unittest.TestCase):
    """Test text cleaning, dictionary scoring, and Excel reporting."""

    def setUp(self) -> None:
        """Initialize test environment and dictionary."""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.base_path = Path(self.temp_dir.name)
        self.ticker = "AAPL"
        self.dict_data = load_loughran_mcdonald_dictionary()

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        self.temp_dir.cleanup()

    def test_dictionary_loading(self) -> None:
        """Verify that all 6 financial sentiment categories are loaded."""
        expected_keys = [
            "positive",
            "negative",
            "uncertainty",
            "litigious",
            "strong_modal",
            "weak_modal",
        ]
        for key in expected_keys:
            self.assertIn(key, self.dict_data)
            self.assertGreater(len(self.dict_data[key]), 0)

        # Check sample canonical words
        self.assertIn("growth", self.dict_data["positive"])
        self.assertIn("loss", self.dict_data["negative"])
        self.assertIn("uncertainty", self.dict_data["uncertainty"])
        self.assertIn("litigation", self.dict_data["litigious"])
        self.assertIn("will", self.dict_data["strong_modal"])
        self.assertIn("could", self.dict_data["weak_modal"])

    def test_clean_mda_text(self) -> None:
        """Verify text cleaning removes punctuation, digits, and stop words."""
        raw_text = (
            "In 2023, our Total Revenues increased 15.5%! However, we faced "
            "severe supply-chain delays and patent litigation in court."
        )
        tokens = clean_mda_text(raw_text)

        # Ensure lowercase
        for token in tokens:
            self.assertTrue(token.islower())
            self.assertTrue(token.isalpha())

        # Ensure stop words are eliminated
        self.assertNotIn("in", tokens)
        self.assertNotIn("by", tokens)
        self.assertNotIn("our", tokens)
        self.assertNotIn("and", tokens)

        # Ensure content words are present
        self.assertIn("revenues", tokens)
        self.assertIn("severe", tokens)
        self.assertIn("litigation", tokens)

    def test_score_mda_sentiment(self) -> None:
        """Verify sentiment metrics, net score, and P/N ratio."""
        # 3 positive ('growth', 'strong', 'profitable')
        # 2 negative ('risk', 'uncertainty')
        # 1 uncertainty ('uncertainty')
        tokens = [
            "growth", "strong", "profitable", "risk", "uncertainty",
            "business", "operations",
        ]
        metrics = score_mda_sentiment(2023, tokens, self.dict_data)

        self.assertEqual(metrics.year, 2023)
        self.assertEqual(metrics.total_words, 7)
        self.assertEqual(metrics.positive_count, 3)
        self.assertEqual(metrics.negative_count, 2)
        self.assertEqual(metrics.net_score, 1)
        self.assertEqual(metrics.pn_ratio, 1.5)
        self.assertEqual(metrics.uncertainty_count, 2)
        self.assertEqual(metrics.tone, "Positive")

    def test_write_sentiment_to_excel_and_verify(self) -> None:
        """Verify creating the Sentiment Analysis tab in Excel."""
        sample_years = [2021, 2022, 2023]
        sample_files = create_sample_mda_files(
            ticker=self.ticker,
            years=sample_years,
            data_dir=self.base_path,
        )
        self.assertEqual(len(sample_files), 3)

        results = process_mda_directory(
            ticker=self.ticker,
            data_dir=self.base_path,
            dictionary=self.dict_data,
        )
        self.assertEqual(len(results), 3)

        excel_path = self.base_path / f"{self.ticker}_Financial_Analysis.xlsx"
        saved_path = write_sentiment_to_excel(
            ticker=self.ticker,
            sentiment_list=results,
            workbook_path=excel_path,
        )
        self.assertTrue(saved_path.is_file())

        wb = openpyxl.load_workbook(saved_path, data_only=True)
        self.assertIn("Sentiment Analysis", wb.sheetnames)

        # Test verification engine
        is_valid = verify_sentiment_tab(saved_path)
        self.assertTrue(is_valid)

    def test_run_sentiment_pipeline(self) -> None:
        """Verify the full end-to-end sentiment analysis pipeline."""
        excel_path = self.base_path / "Pipeline_Test_Analysis.xlsx"
        success = run_sentiment_pipeline(
            ticker=self.ticker,
            data_dir=self.base_path,
            workbook_path=excel_path,
        )
        self.assertTrue(success)
        self.assertTrue(excel_path.is_file())


if __name__ == "__main__":
    unittest.main()
