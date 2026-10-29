"""Generator script to compile the Streamlit financial dashboard into a

standalone, production-grade static HTML single-page application (index.html)
optimized for GitHub Pages hosting.
"""

import json
from pathlib import Path
from typing import Any, Dict

from app import COMPANY_METADATA, load_company_data_bundle
from financial_ratios_analyzer import calculate_pearson_correlation


def extract_full_dataset() -> Dict[str, Any]:
    """Extract and structure audited 10-year datasets for AAPL and MSFT."""
    dataset: Dict[str, Any] = {}

    for ticker in ("AAPL", "MSFT"):
        bundle = load_company_data_bundle(ticker)
        summary = bundle["Executive Summary"]
        prices = [float(p) for p in summary["Stock Price ($)"]]
        sentiment_pn = [float(s) for s in summary["MD&A Sentiment P/N"]]
        corr = calculate_pearson_correlation(prices, sentiment_pn)

        # Clean NaN and None values for clean JSON serialization
        sheets_data: Dict[str, Any] = {}
        for sname, df in bundle.items():
            records = df.to_dict(orient="records")
            cleaned_records = []
            for rec in records:
                cleaned_rec = {}
                for k, v in rec.items():
                    if v is None or (isinstance(v, float) and v != v):
                        cleaned_rec[k] = None
                    else:
                        cleaned_rec[k] = v
                cleaned_records.append(cleaned_rec)
            sheets_data[sname] = cleaned_records

        dataset[ticker] = {
            "metadata": COMPANY_METADATA[ticker],
            "correlation": {
                "r": round(corr, 4),
                "r2": round(corr**2, 4),
            },
            "sheets": sheets_data,
        }

    return dataset


def build_html_file(output_path: Path, template_path: Path) -> None:
    """Generate the complete self-contained index.html file."""
    dataset = extract_full_dataset()
    json_payload = json.dumps(dataset)

    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    html_content = template.replace("__DATA_JSON__", json_payload)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"HTML application built successfully: {output_path.resolve()}")


def main() -> None:
    template_file = Path(__file__).resolve().parent / "dashboard_template.html"
    local_index = Path("index.html")
    build_html_file(local_index, template_file)


if __name__ == "__main__":
    main()
