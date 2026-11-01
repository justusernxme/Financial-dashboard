"""Unit tests for the Streamlit Financial Intelligence Application.

Tests dataset synthesis, company metadata, keyword search and filtering engine,
clean column naming (no unnamed columns), metric rounding, safe float parsing,
and robust handling of raw Excel sheets with footer rows or missing values.
"""

import unittest
import pandas as pd

from app import (
    COMPANY_METADATA,
    clean_ratios_dataframe,
    clean_statement_dataframe,
    clean_summary_dataframe,
    execute_keyword_search,
    format_statement_cell,
    generate_company_dataset,
    safe_float,
)


class TestStreamlitApp(unittest.TestCase):
    """Test suite for Streamlit web app logic and data layer."""

    def setUp(self) -> None:
        """Initialize test parameters and 10-year test datasets."""
        self.years = [2016 + i for i in range(10)]
        self.aapl_data = generate_company_dataset("AAPL", self.years)
        self.msft_data = generate_company_dataset("MSFT", self.years)

    def test_company_metadata(self) -> None:
        """Verify metadata definitions for both AAPL and MSFT."""
        for ticker in ("AAPL", "MSFT"):
            self.assertIn(ticker, COMPANY_METADATA)
            meta = COMPANY_METADATA[ticker]
            self.assertIn("name", meta)
            self.assertIn("icon", meta)
            self.assertIn("sector", meta)
            self.assertIn("industry", meta)
            self.assertIn("exchange", meta)
            self.assertIn("fy_end", meta)

    def test_all_sheets_present(self) -> None:
        """Verify all 7 expected analysis sheets exist for both companies."""
        expected_sheets = [
            "Executive Summary",
            "Income Statement",
            "Balance Sheet",
            "Cash Flows",
            "Financial Ratios",
            "Sentiment Analysis",
            "Stock Price",
        ]
        for sheet in expected_sheets:
            self.assertIn(sheet, self.aapl_data)
            self.assertIn(sheet, self.msft_data)
            self.assertFalse(self.aapl_data[sheet].empty)
            self.assertFalse(self.msft_data[sheet].empty)

    def test_no_unnamed_columns(self) -> None:
        """Verify that zero unnamed columns exist across any DataFrame."""
        for ticker_data in (self.aapl_data, self.msft_data):
            for sheet_name, df in ticker_data.items():
                for col in df.columns:
                    self.assertFalse(
                        str(col).startswith("Unnamed"),
                        f"Found unnamed column '{col}' in '{sheet_name}'",
                    )

    def test_keyword_search_revenue(self) -> None:
        """Verify keyword search finds revenue across multiple sheets."""
        matches = execute_keyword_search(self.aapl_data, "Revenue")
        self.assertFalse(matches.empty)
        # Revenue should match across Income Statement, Summary, Ratios
        matched_items = matches["Metric / Line Item"].str.lower().tolist()
        self.assertTrue(
            any("revenue" in item for item in matched_items),
            "Expected 'revenue' to be found in matched items.",
        )
        self.assertIn("Source Sheet", matches.columns)
        self.assertIn("Metric / Line Item", matches.columns)

    def test_keyword_search_debt(self) -> None:
        """Verify keyword search finds debt metrics."""
        matches = execute_keyword_search(self.msft_data, "Debt")
        self.assertFalse(matches.empty)
        matched_items = matches["Metric / Line Item"].str.lower().tolist()
        self.assertTrue(any("debt" in item for item in matched_items))

    def test_keyword_search_empty_or_nonexistent(self) -> None:
        """Verify search handles empty or unmatched queries gracefully."""
        res_empty = execute_keyword_search(self.aapl_data, "   ")
        self.assertTrue(res_empty.empty)

        res_none = execute_keyword_search(
            self.aapl_data, "XYZNonExistentMetric12345"
        )
        self.assertTrue(res_none.empty)

    def test_summary_rounding_and_types(self) -> None:
        """Verify all numerical values are properly formatted and rounded."""
        summary_df = self.aapl_data["Executive Summary"]
        self.assertEqual(len(summary_df), 10)

        # Check closing stock prices are positive floats
        for price in summary_df["Stock Price ($)"]:
            self.assertIsInstance(price, float)
            self.assertGreater(price, 0.0)

        # Check MD&A Sentiment P/N ratios
        for pn in summary_df["MD&A Sentiment P/N"]:
            self.assertIsInstance(pn, float)
            self.assertGreater(pn, 0.5)

    def test_safe_float_robustness(self) -> None:
        """Verify safe_float handles None, strings with symbols, and NaNs."""
        self.assertEqual(safe_float(None, default=1.0), 1.0)
        self.assertEqual(safe_float(float("nan"), default=0.0), 0.0)
        self.assertEqual(safe_float("$125.50"), 125.50)
        self.assertEqual(safe_float("+15.5%"), 15.5)
        self.assertEqual(safe_float("1.62x"), 1.62)
        self.assertEqual(safe_float("25.0 d"), 25.0)
        self.assertEqual(safe_float("—", default=0.0), 0.0)
        self.assertEqual(safe_float("N/A", default=0.0), 0.0)
        self.assertEqual(safe_float(42), 42.0)
        self.assertEqual(safe_float(-10.25), -10.25)

    def test_clean_summary_dataframe_filters_metadata(self) -> None:
        """Verify footer rows and highlights are filtered from summary."""
        raw_rows = [
            {
                "Fiscal Year": 2022,
                "Stock Price ($)": 150.0,
                "YoY Return": 0.1,
            },
            {
                "Fiscal Year": 2023,
                "Stock Price ($)": 175.0,
                "YoY Return": 0.16,
            },
            {
                "Fiscal Year": "10-YEAR EXECUTIVE HIGHLIGHTS",
                "Stock Price ($)": None,
                "YoY Return": None,
            },
            {
                "Fiscal Year": "Initial Stock Price",
                "Stock Price ($)": "$150.00",
                "YoY Return": None,
            },
        ]
        raw_df = pd.DataFrame(raw_rows)
        cleaned = clean_summary_dataframe(raw_df)

        self.assertEqual(len(cleaned), 2)
        self.assertEqual(cleaned["Fiscal Year"].tolist(), [2022, 2023])
        self.assertIn("YoY Price Return (%)", cleaned.columns)
        # Scaled from 0.10 -> 10.0%
        self.assertAlmostEqual(cleaned["YoY Price Return (%)"].iloc[0], 10.0)

    def test_clean_ratios_dataframe_section_headers(self) -> None:
        """Verify category rows and leading indentation are normalized."""
        raw_rows = [
            {
                "Ratio Category / Metric": "LIQUIDITY RATIOS",
                "Formula / Definition": None,
                "2023": None,
            },
            {
                "Ratio Category / Metric": "    Current Ratio",
                "Formula / Definition": "CA / CL",
                "2023": 1.25,
            },
            {
                "Ratio Category / Metric": "SOLVENCY RATIOS",
                "Formula / Definition": None,
                "2023": None,
            },
            {
                "Ratio Category / Metric": "    Debt-to-Equity",
                "Formula / Definition": "Debt / Eq",
                "2023": 0.85,
            },
        ]
        raw_df = pd.DataFrame(raw_rows)
        cleaned = clean_ratios_dataframe(raw_df)

        self.assertEqual(len(cleaned), 2)
        self.assertEqual(
            cleaned["Ratio / Metric"].tolist(),
            ["Current Ratio", "Debt-to-Equity"],
        )
        self.assertEqual(
            cleaned["Category"].tolist(),
            ["Liquidity", "Solvency"],
        )
        self.assertIn("Format", cleaned.columns)

    def test_all_altair_charts_schema_validation(self) -> None:
        """Verify all Altair charts compile to valid dictionaries."""
        import altair as alt

        # 1. Dual-Axis Line Chart
        sum_df = self.aapl_data["Executive Summary"]
        sum_chart_df = clean_summary_dataframe(sum_df)
        sum_chart_df["Fiscal Year"] = sum_chart_df["Fiscal Year"].astype(str)
        base = alt.Chart(sum_chart_df).encode(
            x=alt.X("Fiscal Year:O", title="Fiscal Year")
        )
        l1 = base.mark_line(color="#1F4E79", strokeWidth=3.5).encode(
            y=alt.Y("Stock Price ($):Q")
        )
        l2 = base.mark_line(color="#C00000", strokeWidth=3.0).encode(
            y=alt.Y("MD&A Sentiment P/N:Q")
        )
        c1 = alt.layer(l1, l2).resolve_scale(y="independent")
        self.assertIsInstance(c1.to_dict(), dict)

        # 2. Sentiment Bar Chart (verifies cornerRadius=4)
        sent_df = self.aapl_data["Sentiment Analysis"]
        sent_chart_df = sent_df.copy()
        sent_chart_df["Fiscal Year"] = sent_chart_df["Fiscal Year"].astype(str)
        c2 = (
            alt.Chart(sent_chart_df)
            .mark_bar(color="#2F5597", cornerRadius=4)
            .encode(
                x=alt.X("Fiscal Year:O"),
                y=alt.Y("P/N Ratio:Q"),
            )
        )
        self.assertIsInstance(c2.to_dict(), dict)

        # 3. Stock Price Area Chart
        stock_df = self.aapl_data["Stock Price"].copy()
        stock_df["Fiscal Year"] = stock_df["Fiscal Year"].astype(str)
        c3 = (
            alt.Chart(stock_df)
            .mark_area(line={"color": "#1F4E79"})
            .encode(
                x=alt.X("Fiscal Year:O"),
                y=alt.Y("Closing Price ($):Q"),
            )
        )
        self.assertIsInstance(c3.to_dict(), dict)

    def test_format_statement_cell(self) -> None:
        """Verify statement cell formatting in millions and em-dash nulls."""
        self.assertEqual(format_statement_cell(None), "—")
        self.assertEqual(format_statement_cell(float("nan")), "—")
        self.assertEqual(format_statement_cell("nan"), "—")
        self.assertEqual(format_statement_cell("None"), "—")
        self.assertEqual(format_statement_cell(""), "—")
        # Raw USD should be scaled to millions
        self.assertEqual(format_statement_cell(93580000000.0), "93,580.00")
        self.assertEqual(format_statement_cell(93580.0), "93,580.00")
        self.assertEqual(format_statement_cell(-15000000000.0), "-15,000.00")
        self.assertEqual(format_statement_cell(2.50), "2.50")
        # Confirm no trailing 'M' or '$'
        formatted = format_statement_cell(93580000000.0)
        self.assertFalse(formatted.endswith("M"))
        self.assertFalse(formatted.startswith("$"))

    def test_clean_statement_dataframe(self) -> None:
        """Verify statement normalization and raw USD scaling."""
        raw_df = pd.DataFrame(
            {
                "Line Item / Financial Concept": [
                    "REVENUES",
                    "Total Net Sales / Revenue",
                    "Operating Expenses",
                ],
                "2015": [None, 93580000000.0, "nan"],
                "2016": ["None", 85320000000.0, ""],
            }
        )
        cleaned = clean_statement_dataframe(raw_df)
        self.assertEqual(cleaned.columns[0], "Line Item")
        self.assertTrue(pd.isna(cleaned.loc[0, "2015"]))
        self.assertEqual(cleaned.loc[1, "2015"], 93580.0)
        self.assertTrue(pd.isna(cleaned.loc[2, "2015"]))
        self.assertEqual(cleaned.loc[1, "2016"], 85320.0)

    def test_statements_formatting_no_nan(self) -> None:
        """Verify all statement sheets have zero nan when formatted."""
        for sheet_name in (
            "Income Statement",
            "Balance Sheet",
            "Cash Flows",
        ):
            df = self.msft_data[sheet_name].copy()
            year_cols = [c for c in df.columns if str(c).isdigit()]
            for yr in year_cols:
                formatted_col = df[yr].apply(format_statement_cell)
                for val in formatted_col:
                    self.assertNotIn(
                        "nan",
                        str(val).lower(),
                        f"Found 'nan' in {sheet_name} column {yr}",
                    )

    def test_goodwill_and_intangibles_present_and_searchable(self) -> None:
        """Verify Goodwill and Intangibles are populated and searchable."""
        # 1. MSFT Balance Sheet check
        msft_bal = self.msft_data["Balance Sheet"]
        line_items = msft_bal["Line Item"].tolist()
        self.assertIn("Goodwill", line_items)
        self.assertIn("Intangible Assets, Net", line_items)

        gw_row = msft_bal[msft_bal["Line Item"] == "Goodwill"].iloc[0]
        # In millions: FY2016 is 17,872.00 and FY2025 is 119,509.00
        self.assertEqual(gw_row["2016"], 17872.0)
        self.assertEqual(gw_row["2025"], 119509.0)

        it_row = msft_bal[
            msft_bal["Line Item"] == "Intangible Assets, Net"
        ].iloc[0]
        self.assertEqual(it_row["2016"], 3733.0)
        self.assertEqual(it_row["2025"], 22604.0)

        # 2. AAPL Balance Sheet check
        aapl_bal = self.aapl_data["Balance Sheet"]
        aapl_items = aapl_bal["Line Item"].tolist()
        self.assertIn("Goodwill", aapl_items)
        self.assertIn("Intangible Assets, Net", aapl_items)

        # 3. Keyword Search check
        search_res = execute_keyword_search(self.msft_data, "Goodwill")
        self.assertFalse(search_res.empty)
        self.assertEqual(
            search_res["Metric / Line Item"].iloc[0], "Goodwill"
        )
        self.assertEqual(search_res["2025"].iloc[0], "119,509.00")

        search_res_it = execute_keyword_search(
            self.msft_data, "Intangible Assets"
        )
        self.assertFalse(search_res_it.empty)
        self.assertEqual(
            search_res_it["Metric / Line Item"].iloc[0],
            "Intangible Assets, Net",
        )
        self.assertEqual(search_res_it["2025"].iloc[0], "22,604.00")

    def test_custom_ticker_support(self) -> None:
        """Verify dynamic custom ticker metadata and 10-year dataset."""
        from app import get_company_metadata, load_company_data_bundle

        # 1. Test common custom ticker metadata
        nvda_meta = get_company_metadata("NVDA")
        self.assertEqual(nvda_meta["name"], "NVIDIA Corporation")
        self.assertIn("Semiconductors", nvda_meta["industry"])

        # 2. Test unknown arbitrary ticker metadata fallback
        xyz_meta = get_company_metadata("XYZ")
        self.assertEqual(xyz_meta["name"], "XYZ Corporation")

        # 3. Test dynamic 10-year dataset loading for custom ticker
        nvda_bundle = load_company_data_bundle("NVDA")
        self.assertIn("Executive Summary", nvda_bundle)
        self.assertIn("Financial Ratios", nvda_bundle)
        self.assertIn("Income Statement", nvda_bundle)
        self.assertIn("Balance Sheet", nvda_bundle)
        self.assertIn("Cash Flows", nvda_bundle)

        nvda_summary = nvda_bundle["Executive Summary"]
        self.assertEqual(len(nvda_summary), 10)
        expected_years = [2016 + i for i in range(10)]
        self.assertEqual(
            nvda_summary["Fiscal Year"].tolist(), expected_years
        )

    def test_company_comparison_pipeline(self) -> None:
        """Verify cross-company comparison data structures and merging."""
        from app import load_company_data_bundle, safe_float

        bundle_a = load_company_data_bundle("AAPL")
        bundle_b = load_company_data_bundle("MSFT")

        sum_a = bundle_a["Executive Summary"]
        sum_b = bundle_b["Executive Summary"]
        self.assertEqual(len(sum_a), 10)
        self.assertEqual(len(sum_b), 10)

        # 1. Test cumulative returns calculation for comparison
        p0_a = safe_float(sum_a.iloc[0]["Stock Price ($)"])
        p10_a = safe_float(sum_a.iloc[-1]["Stock Price ($)"])
        ret_a = ((p10_a - p0_a) / p0_a) * 100.0

        p0_b = safe_float(sum_b.iloc[0]["Stock Price ($)"])
        p10_b = safe_float(sum_b.iloc[-1]["Stock Price ($)"])
        ret_b = ((p10_b - p0_b) / p0_b) * 100.0

        self.assertGreater(ret_a, 0.0)
        self.assertGreater(ret_b, 0.0)

        # 2. Test ratio merging across two companies
        ratios_a = bundle_a["Financial Ratios"]
        ratios_b = bundle_b["Financial Ratios"]

        merged = pd.merge(
            ratios_a,
            ratios_b,
            on=["Ratio / Metric"],
            suffixes=("_AAPL", "_MSFT"),
        )
        self.assertFalse(merged.empty)
        self.assertIn("2025_AAPL", merged.columns)
        self.assertIn("2025_MSFT", merged.columns)

    def test_extract_comparison_metric_val(self) -> None:
        """Verify metric extraction for customizable comparison graph."""
        from app import (
            COMPARISON_METRIC_CONFIG,
            extract_comparison_metric_val,
            load_company_data_bundle,
        )

        bundle = load_company_data_bundle("AAPL")
        self.assertTrue(len(COMPARISON_METRIC_CONFIG) >= 14)

        # Test stock price and revenue
        p_2016 = extract_comparison_metric_val(bundle, "stock_price", 2016)
        p_2025 = extract_comparison_metric_val(bundle, "stock_price", 2025)
        self.assertGreater(p_2025, p_2016)

        rev_2025 = extract_comparison_metric_val(bundle, "revenue", 2025)
        self.assertGreater(rev_2025, 400.0)

        # Test ROE percentage scaling
        roe_2025 = extract_comparison_metric_val(bundle, "roe", 2025)
        self.assertGreater(roe_2025, 50.0)

        # Test Cash Flow metrics
        ocf_2025 = extract_comparison_metric_val(bundle, "ocf", 2025)
        fcf_2025 = extract_comparison_metric_val(
            bundle, "free_cash_flow", 2025
        )
        self.assertGreater(ocf_2025, 100.0)
        self.assertGreater(fcf_2025, 50.0)

        # Test Sentiment metric
        sent_2025 = extract_comparison_metric_val(bundle, "sentiment", 2025)
        self.assertGreater(sent_2025, 0.5)

        # Test Balance Sheet metrics (Total Assets & etc.)
        tot_assets = extract_comparison_metric_val(
            bundle, "total_assets", 2024
        )
        tot_liab = extract_comparison_metric_val(
            bundle, "total_liabilities", 2024
        )
        tot_eq = extract_comparison_metric_val(bundle, "total_equity", 2024)
        cash_eq = extract_comparison_metric_val(
            bundle, "cash_and_equiv", 2024
        )
        self.assertGreater(tot_assets, 300.0)
        self.assertGreater(tot_liab, 200.0)
        self.assertGreater(tot_eq, 40.0)
        self.assertGreater(cash_eq, 15.0)

        # Test Income Statement metrics (EBIT, Gross Profit, EPS)
        ebit_val = extract_comparison_metric_val(bundle, "ebit", 2024)
        gp_val = extract_comparison_metric_val(bundle, "gross_profit", 2024)
        eps_val = extract_comparison_metric_val(bundle, "eps", 2024)
        self.assertGreater(ebit_val, 100.0)
        self.assertGreater(gp_val, 150.0)
        self.assertGreater(eps_val, 5.0)

        # Test Ratios (ROA, Financial Leverage)
        roa_val = extract_comparison_metric_val(bundle, "roa", 2024)
        lev_val = extract_comparison_metric_val(
            bundle, "financial_leverage", 2024
        )
        self.assertGreater(roa_val, 20.0)
        self.assertGreater(lev_val, 3.0)

        # Test custom generated ticker (NVDA)
        nvda_bundle = load_company_data_bundle("NVDA")
        nvda_rev = extract_comparison_metric_val(
            nvda_bundle, "revenue", 2025
        )
        self.assertGreater(nvda_rev, 10.0)
        nvda_assets = extract_comparison_metric_val(
            nvda_bundle, "total_assets", 2025
        )
        self.assertGreater(nvda_assets, 50.0)


if __name__ == "__main__":
    unittest.main()
