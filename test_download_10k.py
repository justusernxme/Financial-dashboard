"""Unit tests for the SEC 10-K Filings Downloader and Reorganizer.

Tests all modular functions including input validation, SEC header parsing,
document selection, file reorganization, and verification.
"""

import tempfile
import unittest
from pathlib import Path

from download_10k import (
    extract_filing_year,
    extract_year_from_accession_number,
    extract_year_from_header_text,
    format_file_size,
    reorganize_filings,
    select_best_document,
    validate_filing_count,
    validate_ticker,
    verify_organized_filings,
)


class TestValidationFunctions(unittest.TestCase):
    """Test user input validation and normalization."""

    def test_validate_ticker_valid(self) -> None:
        """Verify that valid ticker symbols are normalized to uppercase."""
        self.assertEqual(validate_ticker("aapl"), "AAPL")
        self.assertEqual(validate_ticker("  msft  "), "MSFT")
        self.assertEqual(validate_ticker("BRK.A"), "BRK.A")
        self.assertEqual(validate_ticker("brk-b"), "BRK-B")

    def test_validate_ticker_invalid(self) -> None:
        """Verify that invalid ticker symbols raise ValueError."""
        with self.assertRaises(ValueError):
            validate_ticker("")
        with self.assertRaises(ValueError):
            validate_ticker("   ")
        with self.assertRaises(ValueError):
            validate_ticker("AAPL$")
        with self.assertRaises(ValueError):
            validate_ticker("AA PL")

    def test_validate_filing_count_valid(self) -> None:
        """Verify that valid counts are properly parsed as integers."""
        self.assertEqual(validate_filing_count("1"), 1)
        self.assertEqual(validate_filing_count(" 10 "), 10)
        self.assertEqual(validate_filing_count("25"), 25)

    def test_validate_filing_count_invalid(self) -> None:
        """Verify that invalid filing counts raise ValueError."""
        with self.assertRaises(ValueError):
            validate_filing_count("0")
        with self.assertRaises(ValueError):
            validate_filing_count("-5")
        with self.assertRaises(ValueError):
            validate_filing_count("abc")
        with self.assertRaises(ValueError):
            validate_filing_count("")


class TestYearExtraction(unittest.TestCase):
    """Test SEC header parsing and accession number year extraction."""

    def test_extract_year_from_header_period(self) -> None:
        """Verify extraction from CONFORMED PERIOD OF REPORT."""
        header = """
        <SEC-HEADER>0000320193-23-000106.hdr.sgml : 20231103
        CONFORMED SUBMISSION TYPE:	10-K
        CONFORMED PERIOD OF REPORT:	20230930
        FILED AS OF DATE:		20231103
        </SEC-HEADER>
        """
        self.assertEqual(extract_year_from_header_text(header), 2023)

    def test_extract_year_from_header_filed_date_fallback(self) -> None:
        """Verify fallback to FILED AS OF DATE when period is missing."""
        header = """
        <SEC-HEADER>
        CONFORMED SUBMISSION TYPE:	10-K
        FILED AS OF DATE:		20220215
        </SEC-HEADER>
        """
        self.assertEqual(extract_year_from_header_text(header), 2022)

    def test_extract_year_from_accession_number(self) -> None:
        """Verify derivation of year from standard SEC accession numbers."""
        self.assertEqual(
            extract_year_from_accession_number("0000320193-23-000106"), 2023
        )
        self.assertEqual(
            extract_year_from_accession_number("0001193125-15-356351"), 2015
        )
        self.assertEqual(
            extract_year_from_accession_number("0000950123-99-001234"), 1999
        )
        self.assertIsNone(extract_year_from_accession_number("invalid-acc"))

    def test_format_file_size(self) -> None:
        """Verify formatting of byte sizes to human-readable units."""
        self.assertEqual(format_file_size(500), "500.00 B")
        self.assertEqual(format_file_size(2048), "2.00 KB")
        self.assertEqual(format_file_size(1048576 * 3), "3.00 MB")


class TestReorganizationAndVerification(unittest.TestCase):
    """Test file selection, reorganization, and verification end-to-end."""

    def setUp(self) -> None:
        """Create temporary test directories and mock SEC filing tree."""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.base_path = Path(self.temp_dir.name)
        self.raw_base = self.base_path / "raw"
        self.data_base = self.base_path / "data"

        # Create simulated raw SEC EDGAR directory structure
        self.ticker = "AAPL"
        self.acc1_dir = (
            self.raw_base / "sec-edgar-filings" / self.ticker / "10-K"
            / "0000320193-23-000106"
        )
        self.acc2_dir = (
            self.raw_base / "sec-edgar-filings" / self.ticker / "10-K"
            / "0000320193-22-000108"
        )
        self.acc1_dir.mkdir(parents=True, exist_ok=True)
        self.acc2_dir.mkdir(parents=True, exist_ok=True)

        # Mock accession 1: HTML primary document + submission
        header1 = (
            "<SEC-HEADER>\n"
            "CONFORMED SUBMISSION TYPE: 10-K\n"
            "CONFORMED PERIOD OF REPORT: 20230930\n"
            "</SEC-HEADER>\n"
        )
        (self.acc1_dir / "full-submission.txt").write_text(
            header1 + "Mock full submission 2023", encoding="utf-8"
        )
        (self.acc1_dir / "primary-document.html").write_text(
            "<html><body>Apple 2023 10-K Report</body></html>",
            encoding="utf-8",
        )

        # Mock accession 2: Only TXT primary document + submission
        header2 = (
            "<SEC-HEADER>\n"
            "CONFORMED SUBMISSION TYPE: 10-K\n"
            "CONFORMED PERIOD OF REPORT: 20220930\n"
            "</SEC-HEADER>\n"
        )
        (self.acc2_dir / "full-submission.txt").write_text(
            header2 + "Mock full submission 2022", encoding="utf-8"
        )
        (self.acc2_dir / "primary-document.txt").write_text(
            "Apple 2022 10-K Text Report",
            encoding="utf-8",
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        self.temp_dir.cleanup()

    def test_select_best_document(self) -> None:
        """Verify preference for HTML over TXT when available."""
        doc1 = select_best_document(self.acc1_dir)
        self.assertEqual(doc1.name, "primary-document.html")

        doc2 = select_best_document(self.acc2_dir)
        self.assertEqual(doc2.name, "primary-document.txt")

    def test_extract_filing_year(self) -> None:
        """Verify year extraction from accession folder."""
        self.assertEqual(extract_filing_year(self.acc1_dir), 2023)
        self.assertEqual(extract_filing_year(self.acc2_dir), 2022)

    def test_reorganize_and_verify_pipeline(self) -> None:
        """Verify end-to-end reorganization and verification steps."""
        organized = reorganize_filings(
            ticker=self.ticker,
            raw_base_dir=self.raw_base,
            target_base_dir=self.data_base,
        )

        # Expected files: data/AAPL/10K_2022.txt and data/AAPL/10K_2023.html
        self.assertEqual(len(organized), 2)

        file_2023 = self.data_base / self.ticker / "10K_2023.html"
        file_2022 = self.data_base / self.ticker / "10K_2022.txt"

        self.assertTrue(file_2023.is_file())
        self.assertTrue(file_2022.is_file())
        self.assertIn("Apple 2023 10-K Report", file_2023.read_text("utf-8"))
        self.assertIn(
            "Apple 2022 10-K Text Report",
            file_2022.read_text("utf-8"),
        )

        # Test verification function
        result = verify_organized_filings(
            ticker=self.ticker,
            target_base_dir=self.data_base,
            expected_count=2,
        )
        self.assertTrue(result)


if __name__ == "__main__":
    unittest.main()
