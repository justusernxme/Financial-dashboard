"""Unit tests for the SEC 10-K Item 7 (MD&A) Extractor Module.

Tests HTML cleaning, Table of Contents (TOC) bypass heuristics, section
boundary extraction, file saving, and verification.
"""

import tempfile
import unittest
from pathlib import Path

from mda_extractor import (
    clean_html_to_text,
    clean_mda_whitespace,
    extract_mda_from_file,
    extract_mda_from_text,
    extract_year_from_filename,
    format_file_size,
    process_ticker_mda,
    save_mda_text,
    verify_extracted_mda_files,
)


class TestMDAExtractionFunctions(unittest.TestCase):
    """Test core MD&A parsing and text processing algorithms."""

    def test_clean_html_to_text(self) -> None:
        """Verify HTML script/style decomposition and unicode normalization."""
        html = (
            "<html><head><style>body { color: red; }</style></head>"
            "<body><script>alert('test');</script>"
            "<p>Item 7.&nbsp;Management's Discussion</p></body></html>"
        )
        text = clean_html_to_text(html)
        self.assertNotIn("alert", text)
        self.assertNotIn("color: red", text)
        self.assertIn("Item 7. Management's Discussion", text)

    def test_clean_mda_whitespace(self) -> None:
        """Verify normalization of excess newlines and trailing whitespace."""
        raw = "Line 1   \n\n\n\nLine 2   \n\n\nLine 3"
        cleaned = clean_mda_whitespace(raw)
        self.assertEqual(cleaned, "Line 1\n\nLine 2\n\nLine 3")

    def test_extract_year_from_filename(self) -> None:
        """Verify year extraction from standard filing filenames."""
        self.assertEqual(extract_year_from_filename("10K_2023.html"), 2023)
        self.assertEqual(extract_year_from_filename("10K_2022.txt"), 2022)
        self.assertEqual(extract_year_from_filename("10K_2021_2.html"), 2021)
        self.assertIsNone(extract_year_from_filename("random_file.html"))

    def test_format_file_size(self) -> None:
        """Verify byte size formatting."""
        self.assertEqual(format_file_size(100), "100.00 B")
        self.assertEqual(format_file_size(2048), "2.00 KB")
        self.assertEqual(format_file_size(1048576 * 5), "5.00 MB")

    def test_extract_mda_with_toc_bypass(self) -> None:
        """Verify that short Table of Contents entries are bypassed."""
        # Simulated 10-K text with both TOC and substantive body
        text_with_toc = (
            "TABLE OF CONTENTS\n"
            "Item 1. Business ... Page 1\n"
            "Item 7. Management's Discussion and Analysis ... Page 35\n"
            "Item 7A. Quantitative and Qualitative Disclosures ... Page 60\n"
            "Item 8. Financial Statements ... Page 65\n\n"
            "PART II\n"
            "Item 7. Management’s Discussion and Analysis of "
            "Financial Condition and Results of Operations\n\n"
            "Executive Overview:\n"
            "Total revenue for the fiscal year ended September 30, 2023 "
            "was $383.3 billion, an increase of 8%.\n"
            "Liquidity and Capital Resources:\n"
            "As of September 30, 2023, we maintained cash of $29.9B.\n"
            + ("Detailed operational analysis and financial review.\n" * 40)
            + "\nItem 7A. Quantitative and Qualitative Disclosures "
            "About Market Risk\n"
            "Interest rate risk is managed.\n"
            "Item 8. Financial Statements and Supplementary Data\n"
        )

        extracted = extract_mda_from_text(text_with_toc)
        self.assertIsNotNone(extracted)
        # Ensure it captured the substantive body, not the 1-line TOC entry
        self.assertIn("Executive Overview:", extracted)
        self.assertIn("Total revenue for the fiscal year", extracted)
        self.assertIn("Liquidity and Capital Resources:", extracted)
        # Ensure it stopped before Item 7A
        self.assertNotIn("Interest rate risk is managed", extracted)
        self.assertNotIn("Item 8. Financial Statements", extracted)

    def test_extract_mda_ending_at_item_8(self) -> None:
        """Verify boundary detection when Item 7A is omitted."""
        text_no_7a = (
            "Item 7. Management's Discussion and Analysis of "
            "Financial Condition and Results of Operations\n"
            "Fiscal year results showed solid growth across all segments.\n"
            + ("Extensive discussion of operational results.\n" * 40)
            + "\nItem 8. Financial Statements and Supplementary Data\n"
            "Report of independent registered public accounting firm.\n"
        )

        extracted = extract_mda_from_text(text_no_7a)
        self.assertIsNotNone(extracted)
        self.assertIn("Fiscal year results showed solid growth", extracted)
        self.assertNotIn("Report of independent registered", extracted)


class TestMDAPipelineAndVerification(unittest.TestCase):
    """Test full directory processing, saving, and verification."""

    def setUp(self) -> None:
        """Create mock data directory structure."""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.data_dir = Path(self.temp_dir.name)
        self.ticker = "MSFT"
        self.ticker_dir = self.data_dir / self.ticker
        self.ticker_dir.mkdir(parents=True, exist_ok=True)

        # Create mock 10K HTML filing
        html_content = (
            "<html><body>\n"
            "<div><h2>Table of Contents</h2>\n"
            "<p>Item 7. Management's Discussion and Analysis</p>\n"
            "<p>Item 7A. Market Risk</p></div>\n"
            "<div><h2>Item 7. Management’s Discussion and Analysis of "
            "Financial Condition and Results of Operations</h2>\n"
            "<p>Operating Results: Microsoft Cloud revenue up 22%.</p>\n"
            + ("<p>Additional cloud services growth discussion.</p>\n" * 40)
            + "</div>\n"
            "<div><h2>Item 7A. Market Risk</h2></div>\n"
            "</body></html>"
        )
        (self.ticker_dir / "10K_2023.html").write_text(
            html_content, encoding="utf-8"
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        self.temp_dir.cleanup()

    def test_process_ticker_mda_and_verification(self) -> None:
        """Verify processing all filings for a ticker and validating output."""
        mda_files = process_ticker_mda(
            ticker=self.ticker,
            data_dir=self.data_dir,
        )

        # Expected output: data/MSFT/mda/mda_2023.txt
        self.assertEqual(len(mda_files), 1)
        expected_path = self.ticker_dir / "mda" / "mda_2023.txt"
        self.assertEqual(mda_files[0], expected_path)
        self.assertTrue(expected_path.is_file())

        content = expected_path.read_text(encoding="utf-8")
        self.assertIn("Microsoft Cloud revenue up 22%", content)
        self.assertNotIn("Table of Contents", content)

        # Test verification function
        is_valid = verify_extracted_mda_files(mda_files)
        self.assertTrue(is_valid)


if __name__ == "__main__":
    unittest.main()
