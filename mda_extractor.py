#!/usr/bin/env python3
"""SEC 10-K Item 7 (MD&A) Extractor Module.

This module parses downloaded 10-K filings (HTML or TXT) using BeautifulSoup,
identifies and extracts 'Item 7. Management's Discussion and Analysis of
Financial Condition and Results of Operations' (MD&A), and saves each extracted
section to a dedicated text file:

    data/<TICKER>/mda/mda_<YEAR>.txt

It bypasses Table of Contents (TOC) references using narrative length scoring
and structure analysis, verifies extracted file existence and non-zero sizes,
and prints formatted confirmation reports.

Usage:
    Interactive mode:
        python mda_extractor.py

    Command-line argument mode:
        python mda_extractor.py --ticker AAPL
        python mda_extractor.py --input data/AAPL/10K_2023.html
"""

import argparse
import logging
import re
import sys
import unicodedata
from pathlib import Path
from typing import List, Optional, Tuple

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None


# ---------------------------------------------------------------------------
# Constants and Regex Patterns
# ---------------------------------------------------------------------------

DEFAULT_DATA_DIR: Path = Path("data")

# Minimum character count to distinguish true MD&A section from TOC links
MIN_MDA_CHAR_COUNT: int = 1000

# Regex to match Item 7 (MD&A) heading starts.
# Handles: "ITEM 7", "Item 7.", "Item 7 -", non-breaking spaces, curly quotes,
# and optional line break between Item 7 and Management's Discussion.
ITEM_7_START_PATTERN: re.Pattern = re.compile(
    r"(?:^|\n)\s*(?:ITEM|Item)\s+7\s*[\.\:—\-–]?\s*\n?\s*"
    r"Management(?:[’\']s|s)?\s+Discussion\s+and\s+Analysis",
    re.IGNORECASE,
)

# Regex to match Section 7A (Market Risk) or Item 8 (Financial Statements)
# which designate the boundary where MD&A ends.
ITEM_7_END_PATTERN: re.Pattern = re.compile(
    r"(?:^|\n)\s*(?:ITEM|Item)\s+(?:7A|8)\s*[\.\:—\-–]?\s*\n?\s*"
    r"(?:Quantitative\s+and\s+Qualitative\s+Disclosures|"
    r"Financial\s+Statements|\[?Not\s+Applicable\]?|\[?Reserved\]?)?",
    re.IGNORECASE,
)

# Pattern to extract 4-digit year from filing filenames: e.g. 10K_2023.html
FILING_FILENAME_YEAR_REGEX: re.Pattern = re.compile(
    r"^10K_(\d{4})(?:_\d+)?\.(?:html|htm|txt)$", re.IGNORECASE
)


# ---------------------------------------------------------------------------
# Logging and Console Output Helpers
# ---------------------------------------------------------------------------

def setup_logger() -> logging.Logger:
    """Configure and return the application logger.

    Returns:
        logging.Logger: Configured logger instance.
    """
    logger = logging.getLogger("mda_extractor")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%H:%M:%S",
        )
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger


LOGGER: logging.Logger = setup_logger()


def format_file_size(size_bytes: int) -> str:
    """Format byte size into a human-readable string (KB, MB).

    Args:
        size_bytes (int): Size in bytes.

    Returns:
        str: Human-readable string representation of the size.
    """
    for unit in ["B", "KB", "MB", "GB"]:
        if size_bytes < 1024.0 or unit == "GB":
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} B"


# ---------------------------------------------------------------------------
# HTML Parsing and Text Extraction Functions
# ---------------------------------------------------------------------------

def clean_html_to_text(html_content: str) -> str:
    """Parse HTML with BeautifulSoup, remove junk tags, and extract clean text.

    Args:
        html_content (str): Raw HTML content string.

    Returns:
        str: Normalized plain text representation.

    Raises:
        RuntimeError: If BeautifulSoup is not installed.
    """
    if BeautifulSoup is None:
        raise RuntimeError(
            "The 'beautifulsoup4' package is required. "
            "Install it via: pip install beautifulsoup4"
        )

    # Use lxml parser if available for speed; fall back to html.parser
    try:
        soup = BeautifulSoup(html_content, "lxml")
    except Exception:
        soup = BeautifulSoup(html_content, "html.parser")

    # Decompose script, style, and header/footer elements that create noise
    for element in soup(["script", "style", "noscript", "meta"]):
        element.decompose()

    # Extract text with newline separator between blocks
    raw_text = soup.get_text(separator="\n")

    # Normalize unicode (converts \xa0 non-breaking spaces to standard space)
    normalized = unicodedata.normalize("NFKD", raw_text)

    return normalized


def extract_mda_from_text(full_text: str) -> Optional[str]:
    """Identify and extract the Item 7 (MD&A) section from plain text.

    Finds all Item 7 starts and subsequent Item 7A / Item 8 boundaries.
    Filters out short Table of Contents (TOC) snippets and selects the
    substantive discussion narrative.

    Args:
        full_text (str): Normalized text of the 10-K document.

    Returns:
        Optional[str]: Extracted MD&A text if found, else None.
    """
    # Find all occurrences of Item 7 start headings
    start_matches = list(ITEM_7_START_PATTERN.finditer(full_text))
    if not start_matches:
        return None

    # Find all occurrences of section end boundaries (Item 7A or Item 8)
    end_matches = list(ITEM_7_END_PATTERN.finditer(full_text))

    candidates: List[Tuple[int, str]] = []

    for start_match in start_matches:
        start_idx = start_match.start()

        # Find the earliest valid boundary that occurs AFTER this start
        subsequent_ends = [
            em.start() for em in end_matches if em.start() > start_idx
        ]

        if subsequent_ends:
            end_idx = subsequent_ends[0]
            candidate_text = full_text[start_idx:end_idx].strip()
        else:
            # If no explicit Item 7A or 8 was found, slice up to a sane limit
            candidate_text = full_text[start_idx:start_idx + 250000].strip()

        char_len = len(candidate_text)

        # Ignore tiny snippets (almost certainly Table of Contents entries)
        if char_len >= MIN_MDA_CHAR_COUNT:
            candidates.append((char_len, candidate_text))

    if not candidates:
        # If all candidates were below the threshold, fall back to the longest
        all_slices = []
        for start_match in start_matches:
            start_idx = start_match.start()
            subsequent_ends = [
                em.start() for em in end_matches if em.start() > start_idx
            ]
            if subsequent_ends:
                end_idx = subsequent_ends[0]
                all_slices.append(full_text[start_idx:end_idx].strip())
        if all_slices:
            longest = max(all_slices, key=len)
            if len(longest) > 300:
                return clean_mda_whitespace(longest)
        return None

    # Pick the most substantial candidate (true MD&A section)
    candidates.sort(key=lambda item: item[0], reverse=True)
    best_candidate = candidates[0][1]

    return clean_mda_whitespace(best_candidate)


def clean_mda_whitespace(text: str) -> str:
    """Normalize excessive blank lines and whitespace in extracted text.

    Args:
        text (str): Raw extracted text block.

    Returns:
        str: Cleaned text block.
    """
    # Replace sequences of 3 or more newlines with 2 newlines (paragraph break)
    cleaned = re.sub(r"\n{3,}", "\n\n", text)
    # Strip trailing whitespace on each line
    cleaned_lines = [line.rstrip() for line in cleaned.splitlines()]
    return "\n".join(cleaned_lines).strip()


def extract_mda_from_file(file_path: Path) -> Optional[str]:
    """Read a 10-K filing file (HTML or TXT) and extract the MD&A section.

    Args:
        file_path (Path): Path to the 10-K filing file.

    Returns:
        Optional[str]: Extracted MD&A text if successful, else None.
    """
    LOGGER.info("Parsing file: %s...", file_path.name)

    # Attempt reading with UTF-8 first, then Latin-1/CP1252 fallbacks
    content: Optional[str] = None
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            content = file_path.read_text(encoding=enc, errors="replace")
            break
        except Exception:
            continue

    if not content:
        LOGGER.error("Could not read content from file: %s", file_path)
        return None

    # Determine whether file is HTML or plain text
    is_html = file_path.suffix.lower() in (".html", ".htm")
    if is_html:
        clean_text = clean_html_to_text(content)
    else:
        clean_text = unicodedata.normalize("NFKD", content)

    # Extract the MD&A section
    mda_text = extract_mda_from_text(clean_text)
    return mda_text


# ---------------------------------------------------------------------------
# File Organization and Saving
# ---------------------------------------------------------------------------

def extract_year_from_filename(filename: str) -> Optional[int]:
    """Extract the 4-digit year from a standard filing filename.

    Expected filename pattern: 10K_<YEAR>.html or 10K_<YEAR>.txt

    Args:
        filename (str): Name of the filing file.

    Returns:
        Optional[int]: 4-digit integer year, or None if pattern does not match.
    """
    match = FILING_FILENAME_YEAR_REGEX.match(filename)
    if match:
        return int(match.group(1))
    return None


def save_mda_text(mda_text: str, output_path: Path) -> Path:
    """Save extracted MD&A text to the designated destination file.

    Ensures parent directories exist.

    Args:
        mda_text (str): Extracted MD&A text content.
        output_path (Path): Destination file path.

    Returns:
        Path: Path to the saved file.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(mda_text, encoding="utf-8")
    return output_path


def process_ticker_mda(
    ticker: str,
    data_dir: Path = DEFAULT_DATA_DIR,
) -> List[Path]:
    """Extract MD&A sections for all 10-K filings of a given ticker.

    Reads files from `data/<TICKER>/10K_<YEAR>.html` (or .txt) and writes
    extracted sections to `data/<TICKER>/mda/mda_<YEAR>.txt`.

    Args:
        ticker (str): Company stock ticker symbol.
        data_dir (Path): Base data directory containing organized filings.

    Returns:
        List[Path]: List of successfully saved MD&A file paths.
    """
    ticker_dir = data_dir / ticker.upper()
    if not ticker_dir.is_dir():
        LOGGER.error("Ticker directory not found: %s", ticker_dir.resolve())
        return []

    mda_output_dir = ticker_dir / "mda"
    mda_output_dir.mkdir(parents=True, exist_ok=True)

    # Find all 10K filings matching 10K_<YEAR>.*
    filing_files = sorted([
        f for f in ticker_dir.iterdir()
        if f.is_file() and FILING_FILENAME_YEAR_REGEX.match(f.name)
    ])

    if not filing_files:
        LOGGER.warning(
            "No 10-K filing files matching '10K_<YEAR>.*' found in %s.",
            ticker_dir,
        )
        return []

    LOGGER.info(
        "Found %d filing file(s) for %s to extract MD&A from.",
        len(filing_files),
        ticker,
    )

    extracted_files: List[Path] = []

    for filing_file in filing_files:
        year = extract_year_from_filename(filing_file.name)
        if year is None:
            continue

        target_mda_file = mda_output_dir / f"mda_{year}.txt"

        LOGGER.info(
            "Extracting Item 7 (MD&A) for year %d from '%s'...",
            year,
            filing_file.name,
        )

        mda_text = extract_mda_from_file(filing_file)
        if not mda_text:
            LOGGER.warning(
                "Could not locate Item 7 (MD&A) in %s. Skipping.",
                filing_file.name,
            )
            continue

        save_mda_text(mda_text, target_mda_file)
        file_size_str = format_file_size(len(mda_text.encode("utf-8")))
        LOGGER.info(
            "Saved MD&A (%s) -> %s",
            file_size_str,
            target_mda_file.relative_to(data_dir.parent),
        )
        extracted_files.append(target_mda_file)

    return extracted_files


# ---------------------------------------------------------------------------
# Verification Step
# ---------------------------------------------------------------------------

def verify_extracted_mda_files(mda_files: List[Path]) -> bool:
    """Verify that extracted MD&A files exist, are readable, and non-empty.

    Prints a formatted summary table to the console.

    Args:
        mda_files (List[Path]): List of generated MD&A file paths.

    Returns:
        bool: True if all files exist and have non-zero size, False otherwise.
    """
    if not mda_files:
        LOGGER.warning("No MD&A files to verify.")
        return False

    print("\n" + "-" * 75)
    print(" MD&A EXTRACTION VERIFICATION REPORT")
    print("-" * 75)
    header = (
        f"{'Filename':<24} | {'Size':<12} | "
        f"{'Character Count':<18} | {'Status'}"
    )
    print(header)
    print("-" * 75)

    all_valid = True
    verified_count = 0

    for file_path in mda_files:
        if not file_path.is_file():
            row = (
                f"{file_path.name:<24} | {'0 B':<12} | "
                f"{'0':<18} | [FAIL] Missing"
            )
            print(row)
            all_valid = False
            continue

        size_bytes = file_path.stat().st_size
        size_str = format_file_size(size_bytes)

        if size_bytes == 0:
            row = (
                f"{file_path.name:<24} | {size_str:<12} | "
                f"{'0':<18} | [FAIL] Empty"
            )
            print(row)
            all_valid = False
            continue

        # Count character length
        text_content = file_path.read_text(encoding="utf-8", errors="ignore")
        char_count_str = f"{len(text_content):,}"

        row = (
            f"{file_path.name:<24} | {size_str:<12} | "
            f"{char_count_str:<18} | [OK] Verified"
        )
        print(row)
        verified_count += 1

    print("-" * 75)
    print(f"Total verified MD&A files: {verified_count}/{len(mda_files)}")

    return all_valid and verified_count > 0


# ---------------------------------------------------------------------------
# Main Orchestrator and CLI Parser
# ---------------------------------------------------------------------------

def parse_cli_args() -> argparse.Namespace:
    """Parse optional command-line arguments.

    Returns:
        argparse.Namespace: Parsed CLI argument namespace.
    """
    parser = argparse.ArgumentParser(
        description="Extract Item 7 (MD&A) from downloaded 10-K filings.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--ticker",
        type=str,
        help="Stock ticker symbol (e.g. AAPL). If omitted, prompts.",
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=DEFAULT_DATA_DIR,
        help="Base directory containing organized filings.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        help="Direct path to an individual 10-K HTML or TXT file to parse.",
    )

    return parser.parse_args()


def main() -> None:
    """Main execution function for the MD&A extractor."""
    print("=" * 70)
    print("      SEC 10-K ITEM 7 (MD&A) EXTRACTOR (BEAUTIFUL SOUP)")
    print("=" * 70)

    args = parse_cli_args()

    # Case 1: Direct single-file extraction mode
    if args.input:
        input_path = args.input
        if not input_path.is_file():
            LOGGER.error("Specified input file does not exist: %s", input_path)
            sys.exit(1)

        year = extract_year_from_filename(input_path.name) or 2023
        output_dir = input_path.parent / "mda"
        target_file = output_dir / f"mda_{year}.txt"

        mda_text = extract_mda_from_file(input_path)
        if not mda_text:
            LOGGER.error("Failed to extract Item 7 from %s", input_path.name)
            sys.exit(1)

        save_mda_text(mda_text, target_file)
        verify_extracted_mda_files([target_file])
        sys.exit(0)

    # Case 2: Ticker directory processing mode
    if args.ticker:
        ticker = args.ticker.strip().upper()
    else:
        raw_input = input("Enter stock ticker symbol (e.g., AAPL): ").strip()
        if not raw_input:
            LOGGER.error("Ticker symbol cannot be empty.")
            sys.exit(1)
        ticker = raw_input.upper()

    print(f"\nProcessing MD&A extraction for ticker: {ticker}")
    print(f"Source folder: {args.data_dir / ticker}")
    print(f"Output folder: {args.data_dir / ticker / 'mda'}\n")

    extracted_files = process_ticker_mda(ticker=ticker, data_dir=args.data_dir)

    if not extracted_files:
        LOGGER.warning("No MD&A sections could be extracted.")
        sys.exit(1)

    success = verify_extracted_mda_files(extracted_files)
    if success:
        print("\n [SUCCESS] All MD&A sections extracted and verified!")
    else:
        print("\n [WARNING] Verification finished with issues.")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
