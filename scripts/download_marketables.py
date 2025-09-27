"""Download the latest MSPD "Entire" Excel file, extract marketable securities data,
and plot outstanding balances by maturity date.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Tuple

import matplotlib.pyplot as plt
import pandas as pd
import requests

PAGE_DATA_URL = (
    "https://fiscaldata.treasury.gov/page-data/datasets/"
    "monthly-statement-public-debt/page-data.json"
)
BASE_URL = "https://fiscaldata.treasury.gov"
DEFAULT_OUTPUT_DIR = Path("outputs")
DATA_DIR = DEFAULT_OUTPUT_DIR / "data"
FIGURE_DIR = DEFAULT_OUTPUT_DIR / "figures"


@dataclass
class Report:
    report_date: datetime
    path: str

    @property
    def url(self) -> str:
        return BASE_URL + self.path

    @property
    def filename(self) -> str:
        return Path(self.path).name


def fetch_page_data(url: str = PAGE_DATA_URL) -> dict:
    """Fetch the Gatsby page data JSON for the MSPD dataset."""
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return resp.json()


def iter_entire_reports(page_data: dict) -> Iterable[Report]:
    """Yield Report objects for each Excel "Entire" file in the page data."""
    reports = page_data["result"]["pageContext"]["config"]["publishedReports"]
    for report in reports:
        if not report["path"].lower().endswith(".xls"):
            continue
        if "entire" not in report["report_group_desc"].lower():
            continue
        # Drop human readable timezone suffix "(Coordinated ... )"
        date_str = report["report_date"].split(" (")[0]
        # Example: "Sun Aug 31 2025 00:00:00 GMT+0000"
        dt = datetime.strptime(date_str, "%a %b %d %Y %H:%M:%S GMT%z")
        yield Report(report_date=dt, path=report["path"])


def download_file(report: Report, destination: Path) -> Path:
    """Download the Excel file for the provided report."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    resp = requests.get(report.url, timeout=60)
    resp.raise_for_status()
    destination.write_bytes(resp.content)
    return destination


def _detect_cusip_column(df: pd.DataFrame) -> Tuple[str, str]:
    """Return the multi-index column containing CUSIP values."""
    candidates = [col for col in df.columns if col[0] == "Loan Description"]
    best_col = None
    best_matches = 0
    for col in candidates:
        series = df[col].dropna().astype(str).str.upper().str.strip()
        matches = series.str.fullmatch(r"[0-9A-Z]{8,9}").sum()
        if matches > best_matches:
            best_col = col
            best_matches = matches
    if best_col is None or best_matches == 0:
        raise ValueError("Failed to locate CUSIP column in Marketable sheet")
    return best_col


SECURITY_TYPE_MAPPING = {
    "Treasury Bills (Maturity Value):": "T-Bills",
    "Treasury Notes:": "T-Notes",
    "Treasury Bonds:": "T-Bonds",
    "Treasury Inflation-Protected Securities:": "TIPS",
    "Treasury Floating Rate Notes:": "FRNs",
}


def _detect_security_type_column(df: pd.DataFrame) -> Tuple[str, str]:
    """Return the multi-index column containing the security type labels."""

    candidates = [col for col in df.columns if col[0] == "Loan Description"]
    best_col: Tuple[str, str] | None = None
    best_score = 0
    for col in candidates:
        series = df[col].astype(str)
        mask = series.str.contains("Treasury", case=False, na=False)
        score = series[mask].nunique()
        if score > best_score:
            best_col = col
            best_score = score
    if best_col is None or best_score == 0:
        raise ValueError("Failed to locate security type column in Marketable sheet")
    return best_col


def extract_marketable_data(xls_path: Path) -> pd.DataFrame:
    """Load the Marketable sheet and return cleaned marketable security data."""
    df_raw = pd.read_excel(
        xls_path,
        sheet_name="Marketable",
        header=[1, 2],
        engine="xlrd",
    )

    cusip_col = _detect_cusip_column(df_raw)
    security_type_col = _detect_security_type_column(df_raw)
    maturity_col = next(
        (col for col in df_raw.columns if col[0] == "Maturity Date"),
        None,
    )
    if maturity_col is None:
        raise ValueError("Failed to locate Maturity Date column")

    outstanding_col = next(
        (
            col
            for col in df_raw.columns
            if col[0] == "Amount in Millions of Dollars"
            and "Outstanding" in col[1]
        ),
        None,
    )
    if outstanding_col is None:
        raise ValueError("Failed to locate Outstanding column")

    trimmed = df_raw.loc[:, [cusip_col, maturity_col, outstanding_col]].copy()
    trimmed.columns = ["CUSIP", "Maturity Date", "Outstanding"]

    security_types = (
        df_raw[security_type_col]
        .where(
            df_raw[security_type_col]
            .astype(str)
            .str.contains("Treasury", case=False, na=False)
        )
        .ffill()
        .astype(str)
        .str.strip()
        .replace(SECURITY_TYPE_MAPPING)
    )
    trimmed["Security Type"] = security_types

    trimmed["CUSIP"] = (
        trimmed["CUSIP"].astype(str).str.strip().str.upper().replace({"NAN": pd.NA})
    )
    trimmed["Maturity Date"] = pd.to_datetime(
        trimmed["Maturity Date"], errors="coerce"
    )
    trimmed["Outstanding"] = pd.to_numeric(trimmed["Outstanding"], errors="coerce")

    trimmed = trimmed.dropna(
        subset=["CUSIP", "Maturity Date", "Outstanding", "Security Type"]
    )
    trimmed = trimmed[trimmed["CUSIP"].str.fullmatch(r"[0-9A-Z]{8,9}")]
    trimmed = trimmed[trimmed["Outstanding"] > 0]
    trimmed = trimmed[trimmed["Security Type"].isin(SECURITY_TYPE_MAPPING.values())]

    return trimmed.reset_index(drop=True)


def plot_outstanding_by_maturity(data: pd.DataFrame) -> None:
    """Aggregate outstanding balances by maturity month and display a stacked bar chart."""

    type_order = ["T-Notes", "T-Bonds", "TIPS", "FRNs", "T-Bills"]
    colors = {
        "T-Notes": "#1f77b4",
        "T-Bonds": "#ff7f0e",
        "TIPS": "#2ca02c",
        "FRNs": "#d62728",
        "T-Bills": "#9467bd",
    }

    monthly = (
        data.groupby([
            pd.Grouper(key="Maturity Date", freq="MS"),
            "Security Type",
        ])["Outstanding"]
        .sum()
        .unstack(fill_value=0)
        .reindex(columns=type_order, fill_value=0)
        .sort_index()
    )

    fig, ax = plt.subplots(figsize=(14, 6))
    bottom = pd.Series(0, index=monthly.index, dtype=float)
    for security_type in type_order:
        values = monthly[security_type]
        if values.empty:
            continue
        ax.bar(
            monthly.index,
            values,
            bottom=bottom,
            width=20,
            label=security_type,
            color=colors[security_type],
        )
        bottom = bottom + values

    ax.set_title("Marketable Treasury Securities Outstanding by Maturity Month")
    ax.set_xlabel("Maturity Month")
    ax.set_ylabel("Outstanding (Millions of USD)")
    ax.legend(title="Security Type", loc="upper right")
    fig.autofmt_xdate()
    plt.tight_layout()
    plt.show()


def main() -> None:
    page_data = fetch_page_data()
    reports = list(iter_entire_reports(page_data))
    if not reports:
        raise SystemExit("No Excel 'Entire' reports found in page data")

    latest_report = max(reports, key=lambda r: r.report_date)
    print(f"Latest report date: {latest_report.report_date:%Y-%m-%d}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    xls_path = download_file(latest_report, DATA_DIR / latest_report.filename)
    print(f"Downloaded Excel file to {xls_path}")

    marketable = extract_marketable_data(xls_path)
    output_csv = DATA_DIR / "marketable_outstanding.csv"
    marketable.to_csv(output_csv, index=False)
    print(f"Saved cleaned marketable data to {output_csv} ({len(marketable)} rows)")

    plot_outstanding_by_maturity(marketable)


if __name__ == "__main__":
    main()
