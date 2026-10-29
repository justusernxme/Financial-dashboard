# SEC EDGAR 10-K Filings & Financial Analysis Web Application

A robust, modular, and PEP 8 compliant Python analytics suite and interactive
Streamlit web application. The platform downloads 10-K annual filings for US
equities, extracts Item 7 (MD&A), scores text using the Loughran-McDonald
financial sentiment dictionary, retrieves historical balance sheet date stock
prices via `yfinance`, computes 10-year financial ratios across 4 categories,
and exports all data into 7-tab Excel workbooks (`<TICKER>_Financial_Analysis.xlsx`).

An executive-grade Streamlit web application (`app.py`) provides real-time
toggling between Apple Inc. (`AAPL`) and Microsoft Corporation (`MSFT`),
interactive dual-axis visualization, and instant keyword search filtering.

---

## Interactive Streamlit Web Application (`app.py`)

Run the web dashboard with a single command:

```bash
streamlit run app.py
```

### Core Web Dashboard Features:

1. **Seamless Ticker Toggle (`AAPL` vs. `MSFT`)**:
   - Instantaneous switching between Apple Inc. and Microsoft Corporation.
   - Dynamically reloads all company metadata badges, 5 executive KPI cards,
     dual-axis charts, financial statements, ratios, sentiment scores, and
     stock price tables with zero page refresh lag (powered by `@st.cache_data`).

2. **Real-Time Keyword Search & Filter Engine**:
   - Prominently situated search bar allowing live filtering across all
     sheets (e.g., `Revenue`, `Debt`, `Margin`, `ROE`, `Cash`).
   - Matched results display dynamically **above** the rest of the information
     in an executive alert callout with 10-year historical values, source tab
     attribution, and exact metric names.

3. **Executive-Grade Clean Formatting**:
   - **Zero Unnamed Columns**: Completely eliminates artifacts like `Unnamed: 0`.
   - **No Blank Spaces**: Tight, elegant typography without extraneous whitespace.
   - **Strict Decimal Precision**: Currencies formatted as `$#,##0.00`, percentages
     as `0.00%`, multipliers as `0.00x`, days as `0.0 d`, and integers as `#,##0`.

4. **Interactive Dual-Axis Visualization (Altair)**:
   - Primary y-axis: Closing Stock Price ($) in vibrant corporate blue.
   - Secondary y-axis: MD&A Sentiment (P/N Ratio) in emerald green.
   - Hover tooltips showing exact fiscal year, stock price, and P/N sentiment.
   - Pearson correlation coefficient ($r$) badge and statistical interpretation.

5. **Five Central Multi-Tab Explorers**:
   - **Executive Dashboard**: KPI summary cards, dual-axis Altair chart, and
     master cross-sectional 10-year financial table.
   - **Financial Ratios**: 18 financial ratios categorized by Liquidity, Solvency,
     Efficiency, and Profitability with category filter buttons.
   - **Financial Statements**: Audited Income Statement, Balance Sheet, and
     Statement of Cash Flows across all 10 fiscal years.
   - **MD&A Sentiment Analysis**: Loughran-McDonald metrics (Total Words,
     Positive, Negative, Uncertainty, Litigious, Strong/Weak Modal, Net Score,
     Net %, P/N Ratio, Tone badge).
   - **Stock Price History**: Balance sheet date closing prices, closest prior
     trading day resolution, volume, YoY dollar change, and YoY returns.

---

## 7-Tab Consolidated Excel Workbook Structure

Each company's workbook (`<TICKER>_Financial_Analysis.xlsx`) consolidates 7
professionally formatted tabs in executive order:

1. `Executive Summary`: Master cross-sectional table uniting revenues, net income,
   ROE, current ratio, debt-to-equity, stock prices, sentiment P/N ratios,
   dual-axis Excel chart, and Pearson correlation stats.
2. `Financial Ratios`: 18 ratios across Liquidity, Solvency, Efficiency, and
   Profitability side-by-side over 10 years.
3. `Income Statement`: Audited 10-year revenues, expenses, operating and net income.
4. `Balance Sheet`: Audited 10-year assets, liabilities, and stockholders' equity.
5. `Cash Flows`: Audited 10-year operating, investing, and financing cash flows.
6. `Sentiment Analysis`: 10-year Loughran-McDonald metrics scored on Item 7 MD&A.
7. `Stock Price`: Historical closing stock prices and YoY returns on each balance
   sheet date.

---

## Project Structure

```text
AI Project/
├── app.py                              # Streamlit executive web application
├── test_app.py                         # Unit tests for Streamlit application
├── download_10k.py                      # Downloads & reorganizes 10-K filings
├── mda_extractor.py                     # BeautifulSoup parser for Item 7 (MD&A)
├── financial_statements_extractor.py    # 10-year 3-statement Excel extractor
├── sentiment_analyzer.py                # Loughran-McDonald sentiment analysis
├── stock_price_extractor.py             # yfinance stock prices for BS dates
├── financial_ratios_analyzer.py         # 10-year financial ratios & summary chart
├── test_download_10k.py                 # Unit tests for downloading & verification
├── test_mda_extractor.py                # Unit tests for MD&A extraction
├── test_financial_statements_extractor.py # Unit tests for 10-year statements
├── test_sentiment_analyzer.py           # Unit tests for sentiment analysis
├── test_stock_price_extractor.py        # Unit tests for stock price extraction
├── test_financial_ratios_analyzer.py    # Unit tests for ratios & summary dashboard
├── AAPL_Financial_Analysis.xlsx         # Generated 7-tab Excel workbook for Apple
├── MSFT_Financial_Analysis.xlsx         # Generated 7-tab Excel workbook for Microsoft
├── README.md                            # Comprehensive project documentation
└── data/                                # SEC filings and extracted MD&A files
    ├── AAPL/
    │   ├── 10K_2023.html
    │   └── mda/
    │       └── mda_2023.txt
    └── MSFT/
        ├── 10K_2023.html
        └── mda/
            └── mda_2023.txt
```

---

## Installation & Requirements

Ensure Python 3.8+ and the following packages are installed:

```bash
pip install altair beautifulsoup4 lxml openpyxl pandas requests sec-edgar-downloader streamlit yfinance
```

---

## Usage Guide

### 1. Launch the Streamlit Web Application

```bash
streamlit run app.py
```

### 2. Run Financial Ratios & Summary Pipeline (CLI)

```bash
# Calculate 10-year ratios and generate summary dashboard with dual-axis chart
python financial_ratios_analyzer.py --ticker AAPL --years 10
python financial_ratios_analyzer.py --ticker MSFT --years 10
```

### 3. Retrieve Historical Stock Prices for Balance Sheet Dates

```bash
python stock_price_extractor.py --ticker AAPL --years 10
python stock_price_extractor.py --ticker MSFT --years 10
```

### 4. Run Sentiment Analysis on MD&A Text Files

```bash
python sentiment_analyzer.py --ticker AAPL
python sentiment_analyzer.py --ticker MSFT
```

### 5. Extract 10-Year Financial Statements to Excel

```bash
python financial_statements_extractor.py --ticker AAPL --years 10
python financial_statements_extractor.py --ticker MSFT --years 10
```

### 6. Download 10-K Filings & Extract MD&A Sections

```bash
python download_10k.py --ticker AAPL --count 10 --extract-mda
python download_10k.py --ticker MSFT --count 10 --extract-mda
```

---

## Running Unit Tests

Execute the complete test suite:

```bash
python -m unittest discover -p "test_*.py"
```

All unit tests pass cleanly with 100% assertion success and strict PEP 8
compliance ($\le 79$ character line length limit across all codebase files).
# Financial-dashboard
# Financial-dashboard
# Financial-dashboard
# Financial-dashboard
