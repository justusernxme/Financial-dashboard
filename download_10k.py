#!/usr/bin/env python3
"""SEC 10-K Filings Downloader and Reorganizer.

This script prompts the user for a stock ticker symbol (e.g., AAPL) and the
desired number of recent 10-K annual filings to download. It downloads the
filings using either `sec-edgar-downloader` or direct SEC EDGAR REST API
requests (with a compliant User-Agent header complying with SEC fair access
policy).

The script extracts the fiscal reporting year from each filing and reorganizes
the raw files into a clean directory structure:

    data/<TICKER>/10K_<YEAR>.html (or .txt)

It finishes by verifying the reorganized files for existence, proper naming,
and non-zero size, presenting a clear summary table to the console.

Usage:
    Interactive mode:
        python download_10k.py

    Command-line argument mode:
        python download_10k.py --ticker AAPL --count 5
        python download_10k.py --ticker MSFT --count 3 --method sec-api
"""

import argparse
import json
import logging
import re
import shutil
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    from sec_edgar_downloader import Downloader
except ImportError:
    Downloader = None

try:
    from mda_extractor import process_ticker_mda, verify_extracted_mda_files
except ImportError:
    process_ticker_mda = None
    verify_extracted_mda_files = None


# ---------------------------------------------------------------------------
# Constants and SEC EDGAR Configuration
# ---------------------------------------------------------------------------

# Standard SEC EDGAR compliant User-Agent defaults.
# SEC policy requires: "<User-Agent Name> <Email Address>"
DEFAULT_COMPANY_NAME: str = "ResearchAnalyst"
DEFAULT_EMAIL_ADDRESS: str = "research_analyst@example.com"

# Default directory locations
DEFAULT_DATA_DIR: Path = Path("data")
DEFAULT_RAW_DIR: Path = Path("sec-edgar-filings")

# SEC EDGAR REST API endpoints
SEC_TICKERS_URL: str = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL_FMT: str = (
    "https://data.sec.gov/submissions/CIK{cik}.json"
)
SEC_ARCHIVE_URL_FMT: str = (
    "https://www.sec.gov/Archives/edgar/data/{cik}/{acc_no_dash}/{doc}"
)

# Regex pattern for SEC reporting year detection in submission headers
PERIOD_OF_REPORT_REGEX: re.Pattern = re.compile(
    r"(?:CONFORMED\s+)?PERIOD\s+OF\s+REPORT:\s*(\d{4})", re.IGNORECASE
)
FILED_DATE_REGEX: re.Pattern = re.compile(
    r"FILED\s+AS\s+OF\s+DATE:\s*(\d{4})", re.IGNORECASE
)
ACCESSION_YEAR_REGEX: re.Pattern = re.compile(
    r"^\d{10}-(\d{2})-\d{6}$"
)

# Standard filename pattern for organized filings: 10K_<YEAR>.<ext>
ORGANIZED_FILENAME_REGEX: re.Pattern = re.compile(
    r"^10K_(\d{4})(?:_\d+)?\.(html|txt)$", re.IGNORECASE
)


# ---------------------------------------------------------------------------
# Logging and Console Output Helpers
# ---------------------------------------------------------------------------

def setup_logger() -> logging.Logger:
    """Configure and return the application logger.

    Returns:
        logging.Logger: Configured logger instance.
    """
    logger = logging.getLogger("sec_10k_downloader")
    logger.setLevel(logging.INFO)

    # Avoid duplicate handlers if setup_logger is invoked multiple times
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


def print_step_banner(step_num: int, total_steps: int, title: str) -> None:
    """Print a visually distinct header for a processing step.

    Args:
        step_num (int): Current step index (1-based).
        total_steps (int): Total number of steps.
        title (str): Descriptive title of the step.
    """
    separator = "=" * 70
    print(f"\n{separator}")
    print(f" STEP {step_num}/{total_steps}: {title.upper()}")
    print(f"{separator}")


def format_file_size(size_bytes: int) -> str:
    """Format byte size into human-readable string (KB, MB, GB).

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
# Interactive User Input and Validation Functions
# ---------------------------------------------------------------------------

def validate_ticker(ticker_input: str) -> str:
    """Validate and normalize a stock ticker symbol.

    Args:
        ticker_input (str): Raw string entered by the user.

    Returns:
        str: Uppercase, stripped, valid ticker symbol.

    Raises:
        ValueError: If ticker is empty or contains invalid characters.
    """
    cleaned = ticker_input.strip().upper()
    if not cleaned:
        raise ValueError("Ticker symbol cannot be empty.")

    # Allow alphanumeric characters, periods, and hyphens (e.g. BRK.A, BRK-B)
    if not re.match(r"^[A-Z0-9.-]+$", cleaned):
        raise ValueError(
            f"Invalid ticker '{cleaned}'. Tickers may only contain "
            "alphanumeric characters, dots, or dashes."
        )

    return cleaned


def prompt_ticker() -> str:
    """Interactively prompt the user for a valid stock ticker symbol.

    Keeps re-prompting until the user supplies a valid non-empty ticker.

    Returns:
        str: Validated uppercase stock ticker symbol.
    """
    while True:
        try:
            raw_input = input("Enter stock ticker symbol (e.g., AAPL): ")
            validated_ticker = validate_ticker(raw_input)
            print(f" -> Selected ticker: {validated_ticker}")
            return validated_ticker
        except ValueError as err:
            print(f" [!] Input error: {err}. Please try again.\n")


def validate_filing_count(count_input: str) -> int:
    """Validate that the filing count input is a positive integer.

    Args:
        count_input (str): Raw string entered by the user.

    Returns:
        int: Validated positive integer.

    Raises:
        ValueError: If input is not a positive integer greater than zero.
    """
    cleaned = count_input.strip()
    if not cleaned:
        raise ValueError("Filing count cannot be empty.")

    try:
        count = int(cleaned)
    except ValueError as exc:
        raise ValueError(f"'{cleaned}' is not a valid integer.") from exc

    if count <= 0:
        raise ValueError("Number of filings must be greater than zero.")

    return count


def prompt_filing_count() -> int:
    """Interactively prompt for the number of 10-K filings to download.

    Keeps re-prompting until the user enters a valid positive integer.

    Returns:
        int: Validated positive integer.
    """
    while True:
        try:
            prompt_msg = (
                "Enter number of recent 10-K filings to download (e.g., 5): "
            )
            raw_input = input(prompt_msg)
            count = validate_filing_count(raw_input)
            print(f" -> Number of filings requested: {count}")
            return count
        except ValueError as err:
            print(f" [!] Input error: {err}. Please try again.\n")


def prompt_user_agent() -> Tuple[str, str]:
    """Prompt for SEC EDGAR compliant User-Agent headers or use defaults.

    SEC Fair Access Policy requires declared user identification:
    "<Company/Name> <ContactEmail@domain.com>".

    Returns:
        Tuple[str, str]: (company_name, email_address).
    """
    print("\nSEC EDGAR Fair Access Policy requires User-Agent identification.")
    print("Press [Enter] to use standard defaults, or provide custom details.")

    company_prompt = f"Enter organization/name [{DEFAULT_COMPANY_NAME}]: "
    company_input = input(company_prompt).strip()
    company_name = company_input if company_input else DEFAULT_COMPANY_NAME

    email_prompt = f"Enter contact email [{DEFAULT_EMAIL_ADDRESS}]: "
    email_input = input(email_prompt).strip()
    email_address = email_input if email_input else DEFAULT_EMAIL_ADDRESS

    print(f" -> User-Agent declared: {company_name} <{email_address}>")
    return company_name, email_address


# ---------------------------------------------------------------------------
# SEC EDGAR Downloading: Engine 1 (sec-edgar-downloader)
# ---------------------------------------------------------------------------

def download_via_sec_edgar_downloader(
    ticker: str,
    count: int,
    company_name: str,
    email_address: str,
    download_folder: Path,
) -> int:
    """Download 10-K filings using the sec-edgar-downloader library.

    Args:
        ticker (str): Company stock ticker symbol (e.g., 'AAPL').
        count (int): Maximum number of 10-K filings to download.
        company_name (str): Organization name for SEC User-Agent header.
        email_address (str): Contact email for SEC User-Agent header.
        download_folder (Path): Base directory where raw files are saved.

    Returns:
        int: Number of filings successfully downloaded.

    Raises:
        RuntimeError: If sec-edgar-downloader is not installed.
        ValueError: If ticker is unrecognized by SEC EDGAR.
        ConnectionError: If network request to SEC fails.
    """
    if Downloader is None:
        raise RuntimeError(
            "The 'sec-edgar-downloader' library is required but missing. "
            "Please install it using: pip install sec-edgar-downloader"
        )

    LOGGER.info(
        "Initializing sec-edgar-downloader for %s (limit: %d)...",
        ticker,
        count,
    )
    LOGGER.info("Staging raw filings in: %s", download_folder.resolve())

    try:
        downloader = Downloader(
            company_name=company_name,
            email_address=email_address,
            download_folder=download_folder,
        )
    except Exception as exc:
        LOGGER.error("Failed to initialize Downloader: %s", exc)
        raise ConnectionError(
            f"Could not connect to SEC EDGAR API. Details: {exc}"
        ) from exc

    try:
        filings_downloaded = downloader.get(
            "10-K",
            ticker,
            limit=count,
            download_details=True,
        )
        LOGGER.info(
            "sec-edgar-downloader retrieved %d filing(s).",
            filings_downloaded,
        )
        return filings_downloaded

    except ValueError as exc:
        LOGGER.error("SEC Downloader error: %s", exc)
        raise
    except Exception as exc:
        LOGGER.error("Unexpected error during download: %s", exc)
        raise ConnectionError(
            f"An error occurred while downloading filings: {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# SEC EDGAR Downloading: Engine 2 (Direct SEC REST API Requests)
# ---------------------------------------------------------------------------

def fetch_sec_url(url: str, user_agent: str) -> bytes:
    """Fetch content from SEC EDGAR URL with compliant User-Agent.

    Enforces respectful rate-limiting (max 10 requests per second) per SEC
    guidelines.

    Args:
        url (str): Target SEC URL.
        user_agent (str): SEC compliant User-Agent header string.

    Returns:
        bytes: Raw response body.

    Raises:
        ConnectionError: If HTTP request fails.
    """
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        },
    )
    try:
        # Respect SEC rate limits by pausing briefly between requests
        time.sleep(0.12)
        with urllib.request.urlopen(req, timeout=15) as response:
            return response.read()
    except Exception as exc:
        raise ConnectionError(f"HTTP request failed for {url}: {exc}") from exc


def resolve_ticker_to_cik(ticker: str, user_agent: str) -> str:
    """Resolve a stock ticker to a 10-digit zero-padded SEC CIK string.

    Args:
        ticker (str): Company stock ticker symbol (e.g. 'AAPL').
        user_agent (str): Compliant User-Agent header.

    Returns:
        str: 10-digit zero-padded CIK string.

    Raises:
        ValueError: If ticker cannot be resolved to a valid CIK.
    """
    LOGGER.info("Fetching SEC company ticker mapping...")
    data = fetch_sec_url(SEC_TICKERS_URL, user_agent)
    tickers_json = json.loads(data.decode("utf-8"))

    ticker_upper = ticker.upper()
    for item in tickers_json.values():
        if item.get("ticker", "").upper() == ticker_upper:
            cik_int = item.get("cik_str")
            cik_str = str(cik_int).zfill(10)
            LOGGER.info(
                "Resolved ticker '%s' to CIK: %s (%s)",
                ticker,
                cik_str,
                item.get("title", ""),
            )
            return cik_str

    raise ValueError(f"Ticker '{ticker}' was not found in SEC company index.")


def download_via_sec_api(
    ticker: str,
    count: int,
    company_name: str,
    email_address: str,
    download_folder: Path,
) -> int:
    """Download 10-K filings directly using SEC EDGAR REST APIs.

    Args:
        ticker (str): Company stock ticker symbol.
        count (int): Maximum number of 10-K filings to download.
        company_name (str): Organization name for SEC User-Agent header.
        email_address (str): Contact email for SEC User-Agent header.
        download_folder (Path): Staging folder for raw downloads.

    Returns:
        int: Number of filings downloaded.
    """
    user_agent = f"{company_name} {email_address}"
    LOGGER.info(
        "Connecting to SEC EDGAR REST API with User-Agent: %s",
        user_agent,
    )

    cik = resolve_ticker_to_cik(ticker, user_agent)
    submissions_url = SEC_SUBMISSIONS_URL_FMT.format(cik=cik)

    LOGGER.info("Retrieving submission history for CIK %s...", cik)
    submissions_bytes = fetch_sec_url(submissions_url, user_agent)
    submissions = json.loads(submissions_bytes.decode("utf-8"))

    recent = submissions.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    acc_nums = recent.get("accessionNumber", [])
    primary_docs = recent.get("primaryDocument", [])
    filing_dates = recent.get("filingDate", [])
    report_dates = recent.get("reportDate", [])

    downloaded = 0
    cik_no_zeros = cik.lstrip("0")

    for i, form in enumerate(forms):
        if form != "10-K":
            continue

        acc_num = acc_nums[i]
        primary_doc = primary_docs[i]
        filing_date = filing_dates[i]
        report_date = report_dates[i]
        acc_no_dash = acc_num.replace("-", "")

        # Directory structure mirrors sec-edgar-downloader:
        # <download_folder>/sec-edgar-filings/<TICKER>/10-K/<ACCESSION>/
        filing_dir = (
            download_folder / "sec-edgar-filings" / ticker / "10-K" / acc_num
        )
        filing_dir.mkdir(parents=True, exist_ok=True)

        # 1. Download primary document (HTML/XML/TXT)
        doc_filename = Path(primary_doc).name
        # Standardize .htm to .html
        is_html = doc_filename.lower().endswith((".htm", ".html"))
        ext = ".html" if is_html else ".txt"
        primary_target = filing_dir / f"primary-document{ext}"

        if not primary_target.is_file():
            primary_url = SEC_ARCHIVE_URL_FMT.format(
                cik=cik_no_zeros,
                acc_no_dash=acc_no_dash,
                doc=doc_filename,
            )
            LOGGER.info(
                "Downloading 10-K (%s, period: %s) -> %s",
                acc_num,
                report_date,
                primary_target.name,
            )
            content = fetch_sec_url(primary_url, user_agent)
            primary_target.write_bytes(content)

        # 2. Write submission header for robust fiscal year extraction
        full_sub_file = filing_dir / "full-submission.txt"
        if not full_sub_file.is_file():
            header_content = (
                f"<SEC-HEADER>\n"
                f"ACCESSION NUMBER: {acc_num}\n"
                f"CONFORMED SUBMISSION TYPE: 10-K\n"
                f"CONFORMED PERIOD OF REPORT: {report_date.replace('-', '')}\n"
                f"FILED AS OF DATE: {filing_date.replace('-', '')}\n"
                f"</SEC-HEADER>\n"
            )
            full_sub_file.write_text(header_content, encoding="utf-8")

        downloaded += 1
        if downloaded >= count:
            break

    LOGGER.info("Direct SEC API download complete: %d retrieved.", downloaded)
    return downloaded


def download_10k_filings(
    ticker: str,
    count: int,
    company_name: str,
    email_address: str,
    download_folder: Path,
    method: str = "auto",
) -> int:
    """Download 10-K filings using either sec-edgar-downloader or SEC REST API.

    Args:
        ticker (str): Company stock ticker symbol.
        count (int): Maximum number of 10-K filings to download.
        company_name (str): Organization name for SEC User-Agent header.
        email_address (str): Contact email for SEC User-Agent header.
        download_folder (Path): Base directory where raw files are saved.
        method (str): 'auto', 'downloader', or 'sec-api'.

    Returns:
        int: Number of filings downloaded.
    """
    if method == "downloader" or (method == "auto" and Downloader is not None):
        try:
            return download_via_sec_edgar_downloader(
                ticker=ticker,
                count=count,
                company_name=company_name,
                email_address=email_address,
                download_folder=download_folder,
            )
        except Exception as exc:
            if method == "auto":
                LOGGER.warning(
                    "sec-edgar-downloader failed (%s). Using SEC API...",
                    exc,
                )
            else:
                raise

    # Fallback or explicit direct SEC API method
    return download_via_sec_api(
        ticker=ticker,
        count=count,
        company_name=company_name,
        email_address=email_address,
        download_folder=download_folder,
    )


# ---------------------------------------------------------------------------
# Header Parsing and Year Extraction
# ---------------------------------------------------------------------------

def extract_year_from_header_text(header_text: str) -> Optional[int]:
    """Parse text from SEC filing header to extract fiscal reporting year.

    Checks CONFORMED PERIOD OF REPORT first (fiscal year), then falls back
    to FILED AS OF DATE (filing year).

    Args:
        header_text (str): String content from the SEC header section.

    Returns:
        Optional[int]: 4-digit integer year if found, else None.
    """
    # 1. Primary choice: Reporting period date (CONFORMED PERIOD OF REPORT)
    match_period = PERIOD_OF_REPORT_REGEX.search(header_text)
    if match_period:
        year = int(match_period.group(1))
        if 1900 <= year <= 2100:
            return year

    # 2. Secondary fallback: Filing date (FILED AS OF DATE)
    match_filed = FILED_DATE_REGEX.search(header_text)
    if match_filed:
        year = int(match_filed.group(1))
        if 1900 <= year <= 2100:
            return year

    return None


def extract_year_from_accession_number(accession_num: str) -> Optional[int]:
    """Derive filing year from standard 20-character SEC accession number.

    Standard accession number format: CIK(10)-YY-Sequence(6), e.g.
    '0000320193-23-000106' where '23' corresponds to 2023.

    Args:
        accession_num (str): SEC accession number string.

    Returns:
        Optional[int]: 4-digit integer year, or None if pattern does not match.
    """
    match = ACCESSION_YEAR_REGEX.match(accession_num.strip())
    if not match:
        return None

    two_digit_year = int(match.group(1))
    # SEC EDGAR filings began electronically in 1994
    if two_digit_year >= 90:
        return 1900 + two_digit_year
    return 2000 + two_digit_year


def extract_filing_year(accession_dir: Path) -> int:
    """Determine the filing year for an accession folder.

    Reads the SEC header lines from 'full-submission.txt'. If not found or
    unreadable, falls back to the accession number's 2-digit year code.

    Args:
        accession_dir (Path): Directory containing filing files.

    Returns:
        int: 4-digit filing year.

    Raises:
        ValueError: If no valid year can be identified.
    """
    accession_name = accession_dir.name
    full_submission_file = accession_dir / "full-submission.txt"

    if full_submission_file.is_file():
        try:
            # Read first 500 lines to avoid loading huge multi-MB files
            header_lines = []
            with open(
                full_submission_file,
                mode="r",
                encoding="utf-8",
                errors="ignore",
            ) as file_handle:
                for _ in range(500):
                    line = file_handle.readline()
                    if not line:
                        break
                    header_lines.append(line)
                    if "</SEC-HEADER>" in line:
                        break

            header_text = "".join(header_lines)
            year = extract_year_from_header_text(header_text)
            if year is not None:
                return year

        except OSError as exc:
            LOGGER.warning(
                "Could not read full-submission.txt in %s: %s",
                accession_dir,
                exc,
            )

    # Fallback to deriving year from accession number
    fallback_year = extract_year_from_accession_number(accession_name)
    if fallback_year is not None:
        LOGGER.info(
            "Extracted year %d from accession number %s (fallback)",
            fallback_year,
            accession_name,
        )
        return fallback_year

    raise ValueError(
        f"Unable to determine filing year for accession directory: "
        f"{accession_dir}"
    )


# ---------------------------------------------------------------------------
# Document Selection and File Organization
# ---------------------------------------------------------------------------

def select_best_document(accession_dir: Path) -> Path:
    """Identify the primary filing document inside an accession folder.

    Prefers HTML documents (such as 'primary-document.html' or any .htm/.html
    excluding raw full-submission), then text documents (.txt).

    Args:
        accession_dir (Path): Directory containing downloaded filing files.

    Returns:
        Path: Path to the selected primary document.

    Raises:
        FileNotFoundError: If accession directory has no candidate document.
    """
    candidates = list(accession_dir.iterdir())

    # 1. Look for explicit primary-document.html or .htm
    for candidate in candidates:
        cand_lower = candidate.name.lower()
        if cand_lower in ("primary-document.html", "primary-document.htm"):
            return candidate

    # 2. Look for any HTML file in the folder
    for candidate in candidates:
        if (
            candidate.suffix.lower() in (".html", ".htm")
            and candidate.is_file()
        ):
            return candidate

    # 3. Look for explicit primary-document.txt
    for candidate in candidates:
        if (
            candidate.name.lower() == "primary-document.txt"
            and candidate.is_file()
        ):
            return candidate

    # 4. Fall back to full-submission.txt
    full_submission = accession_dir / "full-submission.txt"
    if full_submission.is_file():
        return full_submission

    # 5. Fall back to any text file
    for candidate in candidates:
        if candidate.suffix.lower() == ".txt" and candidate.is_file():
            return candidate

    raise FileNotFoundError(
        f"No valid filing documents found inside: {accession_dir}"
    )


def reorganize_filings(
    ticker: str,
    raw_base_dir: Path,
    target_base_dir: Path,
) -> List[Path]:
    """Reorganize raw downloaded filings into clean destination structure:
    `data/<TICKER>/10K_<YEAR>.html` (or .txt).

    Args:
        ticker (str): Stock ticker symbol (e.g. 'AAPL').
        raw_base_dir (Path): Path where raw files were downloaded.
        target_base_dir (Path): Destination base directory (e.g., 'data').

    Returns:
        List[Path]: List of newly created/organized file paths.
    """
    possible_raw_dirs = [
        raw_base_dir / "sec-edgar-filings" / ticker / "10-K",
        raw_base_dir / ticker / "10-K",
    ]

    source_dir: Optional[Path] = None
    for candidate_dir in possible_raw_dirs:
        if candidate_dir.is_dir():
            source_dir = candidate_dir
            break

    if source_dir is None:
        LOGGER.warning(
            "Raw filings directory not found for %s under %s.",
            ticker,
            raw_base_dir,
        )
        return []

    # Prepare destination directory: data/<TICKER>/
    destination_dir = target_base_dir / ticker
    destination_dir.mkdir(parents=True, exist_ok=True)
    LOGGER.info("Destination directory ready: %s", destination_dir.resolve())

    # Find all accession subdirectories (e.g. 0000320193-23-000106)
    accession_dirs = [p for p in source_dir.iterdir() if p.is_dir()]
    LOGGER.info(
        "Discovered %d raw accession folders to process.",
        len(accession_dirs),
    )

    organized_files: List[Path] = []
    seen_target_names: Dict[str, int] = {}

    for accession_dir in sorted(accession_dirs):
        try:
            # 1. Determine filing year
            year = extract_filing_year(accession_dir)

            # 2. Select best document (HTML or TXT)
            best_doc = select_best_document(accession_dir)
            if best_doc.suffix.lower() in (".html", ".htm"):
                ext = ".html"
            else:
                ext = ".txt"

            # 3. Determine unique target filename: 10K_<YEAR>.<ext>
            base_name = f"10K_{year}"
            count_for_year = seen_target_names.get(base_name, 0)
            seen_target_names[base_name] = count_for_year + 1

            if count_for_year == 0:
                target_filename = f"{base_name}{ext}"
            else:
                # Disambiguate if multiple filings share same year (amends)
                target_filename = f"{base_name}_{count_for_year + 1}{ext}"

            target_path = destination_dir / target_filename

            # 4. Copy file with preserved metadata
            LOGGER.info(
                "Copying '%s' -> '%s'...",
                best_doc.name,
                target_path.relative_to(target_base_dir.parent),
            )
            shutil.copy2(best_doc, target_path)
            organized_files.append(target_path)

        except Exception as exc:
            LOGGER.error(
                "Failed to reorganize filing from %s: %s",
                accession_dir.name,
                exc,
            )

    return organized_files


# ---------------------------------------------------------------------------
# Verification Step
# ---------------------------------------------------------------------------

def verify_organized_filings(
    ticker: str,
    target_base_dir: Path,
    expected_count: Optional[int] = None,
) -> bool:
    """Verify that reorganized files exist, are properly named, and non-empty.

    Prints a formatted summary table to the console.

    Args:
        ticker (str): Stock ticker symbol.
        target_base_dir (Path): Destination base directory.
        expected_count (Optional[int]): Expected number of files, if known.

    Returns:
        bool: True if verification succeeds without errors, False otherwise.
    """
    destination_dir = target_base_dir / ticker

    if not destination_dir.is_dir():
        LOGGER.error(
            "Verification failed: Destination directory not found: %s",
            destination_dir,
        )
        return False

    files = sorted(destination_dir.iterdir())
    if not files:
        LOGGER.warning(
            "Verification: No files found in directory %s.",
            destination_dir,
        )
        return False

    print("\n" + "-" * 75)
    print(f" VERIFICATION REPORT: {destination_dir.resolve()}")
    print("-" * 75)
    header = (
        f"{'Filename':<22} | {'Year':<6} | {'Format':<6} | "
        f"{'Size':<12} | {'Status'}"
    )
    print(header)
    print("-" * 75)

    all_valid = True
    verified_count = 0

    for file_path in files:
        if not file_path.is_file():
            continue

        filename = file_path.name
        match = ORGANIZED_FILENAME_REGEX.match(filename)
        size_bytes = file_path.stat().st_size
        size_str = format_file_size(size_bytes)

        if not match:
            row = (
                f"{filename:<22} | {'N/A':<6} | {'N/A':<6} | "
                f"{size_str:<12} | [FAIL] Invalid Name"
            )
            print(row)
            all_valid = False
            continue

        year_str = match.group(1)
        ext_str = match.group(2).lower()

        if size_bytes == 0:
            row = (
                f"{filename:<22} | {year_str:<6} | {ext_str:<6} | "
                f"{size_str:<12} | [FAIL] Empty File"
            )
            print(row)
            all_valid = False
            continue

        row = (
            f"{filename:<22} | {year_str:<6} | {ext_str:<6} | "
            f"{size_str:<12} | [OK] Verified"
        )
        print(row)
        verified_count += 1

    print("-" * 75)
    print(
        f"Total verified files: {verified_count} "
        + (f"(Expected: {expected_count})" if expected_count else "")
    )

    if expected_count is not None and verified_count < expected_count:
        LOGGER.warning(
            "Note: %d filings were verified out of %d requested "
            "(older filings may not be present on SEC EDGAR for this entity).",
            verified_count,
            expected_count,
        )

    return all_valid and verified_count > 0


# ---------------------------------------------------------------------------
# Main Orchestrator and CLI Parser
# ---------------------------------------------------------------------------

def parse_cli_args() -> argparse.Namespace:
    """Parse optional command-line arguments for automated or script usage.

    Returns:
        argparse.Namespace: Parsed CLI argument namespace.
    """
    parser = argparse.ArgumentParser(
        description="Download and re-organize SEC EDGAR 10-K filings.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--ticker",
        type=str,
        help="Stock ticker symbol (e.g. AAPL). If omitted, prompts.",
    )
    parser.add_argument(
        "--count",
        type=int,
        help="Number of recent 10-K filings to download. If omitted, prompts.",
    )
    parser.add_argument(
        "--company",
        type=str,
        default=DEFAULT_COMPANY_NAME,
        help="Organization or name for SEC EDGAR User-Agent compliance.",
    )
    parser.add_argument(
        "--email",
        type=str,
        default=DEFAULT_EMAIL_ADDRESS,
        help="Contact email address for SEC EDGAR User-Agent compliance.",
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=DEFAULT_DATA_DIR,
        help="Target base directory for organized filings.",
    )
    parser.add_argument(
        "--raw-dir",
        type=Path,
        default=DEFAULT_RAW_DIR,
        help="Directory to stage raw SEC downloads.",
    )
    parser.add_argument(
        "--method",
        type=str,
        choices=["auto", "downloader", "sec-api"],
        default="auto",
        help="Download backend: 'downloader', 'sec-api', or 'auto'.",
    )
    parser.add_argument(
        "--extract-mda",
        action="store_true",
        help="Extract Item 7 (MD&A) sections into data/<TICKER>/mda/.",
    )

    return parser.parse_args()


def run_pipeline(
    ticker: str,
    count: int,
    company_name: str,
    email_address: str,
    data_dir: Path,
    raw_dir: Path,
    method: str = "auto",
    extract_mda: bool = False,
) -> bool:
    """Execute end-to-end 10-K download, reorganization, and verification.

    Args:
        ticker (str): Stock ticker symbol.
        count (int): Number of filings to download.
        company_name (str): SEC User-Agent company name.
        email_address (str): SEC User-Agent email address.
        data_dir (Path): Destination base directory.
        raw_dir (Path): Raw staging directory.
        method (str): Download backend.
        extract_mda (bool): Whether to extract Item 7 (MD&A) sections.

    Returns:
        bool: True if process completes successfully, False otherwise.
    """
    total_steps = 4 if extract_mda else 3

    # Step 1: Download Filings via SEC EDGAR
    print_step_banner(
        step_num=1,
        total_steps=total_steps,
        title=f"Downloading {count} 10-K filings for {ticker}",
    )
    try:
        download_count = download_10k_filings(
            ticker=ticker,
            count=count,
            company_name=company_name,
            email_address=email_address,
            download_folder=raw_dir,
            method=method,
        )
    except Exception as exc:
        LOGGER.error("Step 1 failed: %s", exc)
        return False

    if download_count == 0:
        LOGGER.warning(
            "No filings downloaded. Check if '%s' has 10-Ks on EDGAR.",
            ticker,
        )
        return False

    # Step 2: Reorganize and rename files into data/<TICKER>/10K_<YEAR>.*
    print_step_banner(
        step_num=2,
        total_steps=total_steps,
        title=f"Reorganizing and renaming files into {data_dir}/{ticker}/",
    )
    try:
        organized_files = reorganize_filings(
            ticker=ticker,
            raw_base_dir=raw_dir,
            target_base_dir=data_dir,
        )
        LOGGER.info(
            "Successfully organized %d filing(s) into %s/%s/",
            len(organized_files),
            data_dir,
            ticker,
        )
    except Exception as exc:
        LOGGER.error("Step 2 failed during file reorganization: %s", exc)
        return False

    # Step 3: Verify reorganized files
    print_step_banner(
        step_num=3,
        total_steps=total_steps,
        title="Verifying file existence, naming, and integrity",
    )
    verification_success = verify_organized_filings(
        ticker=ticker,
        target_base_dir=data_dir,
        expected_count=download_count,
    )

    if not verification_success:
        print("\n [WARNING] Verification finished with issues. Check logs.")
        return False

    # Step 4 (Optional): Extract Item 7 (MD&A)
    if extract_mda:
        if process_ticker_mda is not None:
            print_step_banner(
                step_num=4,
                total_steps=total_steps,
                title=f"Extracting Item 7 (MD&A) for {ticker}",
            )
            mda_files = process_ticker_mda(ticker=ticker, data_dir=data_dir)
            if verify_extracted_mda_files is not None:
                verify_extracted_mda_files(mda_files)
        else:
            LOGGER.warning("mda_extractor module not available.")

    print("\n [SUCCESS] Filings downloaded, organized, and verified!")
    return True


def main() -> None:
    """Main execution function for the SEC 10-K downloader tool."""
    print("=" * 70)
    print("        SEC EDGAR 10-K FILINGS DOWNLOADER & ORGANIZER")
    print("=" * 70)

    args = parse_cli_args()

    # Determine ticker: Use CLI argument if provided, else prompt
    if args.ticker:
        try:
            ticker = validate_ticker(args.ticker)
            print(f"Stock ticker provided via CLI: {ticker}")
        except ValueError as err:
            LOGGER.error("CLI argument error: %s", err)
            sys.exit(1)
    else:
        ticker = prompt_ticker()

    # Determine count: Use CLI argument if provided, else prompt
    if args.count is not None:
        try:
            count = validate_filing_count(str(args.count))
            print(f"Filing count provided via CLI: {count}")
        except ValueError as err:
            LOGGER.error("CLI argument error: %s", err)
            sys.exit(1)
    else:
        count = prompt_filing_count()

    # Prompt or use User-Agent credentials
    company_name = args.company
    email_address = args.email

    # If running purely interactively, confirm or prompt User-Agent
    if not args.ticker and args.count is None:
        company_name, email_address = prompt_user_agent()

    # Check if MD&A extraction is requested or prompt if interactive
    extract_mda = args.extract_mda
    if not args.ticker and args.count is None and not args.extract_mda:
        mda_choice = input(
            "\nExtract Item 7 (MD&A) sections to text files? [Y/n]: "
        ).strip().lower()
        extract_mda = mda_choice in ("", "y", "yes")

    # Run the processing pipeline
    success = run_pipeline(
        ticker=ticker,
        count=count,
        company_name=company_name,
        email_address=email_address,
        data_dir=args.data_dir,
        raw_dir=args.raw_dir,
        method=args.method,
        extract_mda=extract_mda,
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
