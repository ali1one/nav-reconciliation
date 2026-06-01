# NAV Reconciliation — Step 1: Data Ingestion & Coverage Check

## Overview

This tool reads NAV (Net Asset Value) data from three independent sources — an internal system, a Fund Administrator, and a Custodian — merges them on ISIN, and produces a coverage matrix that flags every fund missing from any source.

It is designed to run daily before the break computation step, ensuring no fund is reconciled against incomplete data.

---

## Project Structure

```
Step1/
├── nav_coverage_check.py       # Main script
├── data/
│   ├── internal/               # Internal system exports (.xlsx)
│   ├── admin/                  # Fund Administrator reports (.csv)
│   └── custodian/              # Custodian valuations (.csv)
└── output/                     # Generated Excel reports
```

---

## Input Files

| Source | Folder | Format | Example filename |
|---|---|---|---|
| Internal system | `data/internal/` | Excel `.xlsx` | `internal_nav_20260530.xlsx` |
| Fund Administrator | `data/admin/` | CSV | `admin_nav_report_20260530.csv` |
| Custodian | `data/custodian/` | CSV | `custodian_valuations_20260530.csv` |

The script always picks the most recent file in each folder, so just drop the new file in — no config change needed.

---

## Output

An Excel file is written to `output/` named `step1_coverage_YYYY-MM-DD.xlsx` with two sheets:

- **Coverage Matrix** — one row per ISIN, with a ✓/✗ flag for each source
- **Missing Funds** — only the incomplete ISINs, with a category and action note

### Missing Fund Categories

| Category | Meaning | Action |
|---|---|---|
| `LATE_FILE` | Present in ≥2 sources, absent from 1 | Chase counterparty for resubmission |
| `NOT_ONBOARDED` | At counterparties but not in internal system | Verify onboarding with Fund Operations |
| `NEW_FUND_SETUP` | In internal but absent from counterparties | Confirm counterparty setup is in progress |
| `ORPHAN` | Found in only 1 source | Likely a data error — investigate before reconciling |

---

## How to Run

### 1. Install dependencies

```bash
pip install pandas openpyxl
```

### 2. Drop input files into the correct folders

```
data/internal/    ← internal_nav_YYYYMMDD.xlsx
data/admin/       ← admin_nav_report_YYYYMMDD.csv
data/custodian/   ← custodian_valuations_YYYYMMDD.csv
```

### 3. Run the script

```bash
python nav_coverage_check.py
```

A summary is printed to the console and the Excel report is saved to `output/`.

---

## Dependencies

- Python 3.10+
- `pandas`
- `openpyxl`

---

## Author

Youssef Ait Ali
