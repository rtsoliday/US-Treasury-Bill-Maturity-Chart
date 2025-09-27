# US-Treasury-Bill-Maturity-Chart

## Overview

This repository now includes a utility script that retrieves the most recent
"Entire" Excel workbook from the U.S. Treasury Fiscal Data Monthly Statement of
Public Debt (MSPD) dataset, extracts marketable securities information, and
generates a bar chart showing the total outstanding balance by maturity date.

## Prerequisites

Install the required Python packages (Python 3.10+ recommended):

```bash
pip install pandas matplotlib requests xlrd python-dateutil
```

## Usage

Run the script from the repository root:

```bash
python scripts/download_marketables.py
```

To download a specific month's report instead of the latest one, provide the
`--report-date` option using a `YYYY-MM` value:

```bash
python scripts/download_marketables.py --report-date 2024-05
```

The script will:

1. Download the requested (latest by default) MSPD `Entire` Excel file to
   `outputs/data/`.
2. Extract the CUSIP, maturity date, and outstanding balance (in millions of
   dollars) for marketable securities, saving the cleaned dataset to
   `outputs/data/marketable_outstanding.csv`.
3. Aggregate outstanding balances by maturity date and render a bar chart saved
   to `outputs/figures/marketable_outstanding_by_maturity.png`.

Both the CSV and the chart are ignored by Git so you can regenerate them at any time.

