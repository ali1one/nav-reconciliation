# NAV Reconciliation — Internal vs Administrator vs Custodian

Daily NAV control for a range of UCITS funds, in two steps:

| Step | Script | What it does |
|---|---|---|
| 1 | `nav_coverage_check.py` | Loads the three sources, checks every fund is present everywhere |
| 2 | `nav_break_computation.py` | Measures NAV / unit differences in bps, flags breaks, finds the probable faulty source |

---

## Step 1 — Data Ingestion & Coverage Check

This tool reads NAV (Net Asset Value) data from three independent sources — an internal system, a Fund Administrator, and a Custodian — merges them on ISIN, and produces a coverage matrix that flags every fund missing from any source.

It is designed to run daily before the break computation step, ensuring no fund is reconciled against incomplete data.

---

## Project Structure

```
nav-reconciliation/
├── nav_coverage_check.py       # Step 1 — coverage check
├── nav_break_computation.py    # Step 2 — break computation
├── data/
│   ├── internal/               # Internal system exports (.xlsx)
│   ├── admin/                  # Fund Administrator reports (.csv)
│   └── custodian/              # Custodian valuations (.csv)
├── samples/breaks_demo/        # Same data with 5 injected anomalies (demo)
├── tests/                      # pytest suite for Step 2
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

## Step 2 — Break Computation

For every fund present in all three sources, Internal is compared with the Administrator and with the Custodian:

| Metric | Formula |
|---|---|
| NAV per share difference | `(NAVPS counterparty / NAVPS internal − 1) × 10 000` bps |
| Units difference | `(Units counterparty / Units internal − 1) × 10 000` bps |
| Total NAV difference | `TotalNAV counterparty − TotalNAV internal` (fund currency) |

### Tolerances (configurable)

| Status | Rule (defaults) |
|---|---|
| `MATCHED` | NAVPS ≤ 1 bp and units ≤ 1 bp |
| `WARNING` | 1 bp < NAVPS ≤ 5 bps |
| `BREAK` | NAVPS > 5 bps, units > 1 bp, different NAV date or static data |

Per-fund overrides live in `TOLERANCE_OVERRIDES` (the money market fund uses a tighter 0.5 / 2 bps band).

### Root cause & triangulation

| Root cause | Meaning |
|---|---|
| `PRICING` | NAV per share differs, unit count agrees → price / valuation issue |
| `CAPITAL_ACTIVITY` | Unit count differs → subscription / redemption not booked |
| `NAV_DATE` | Sources valued on different dates |
| `STATIC_DATA` | Currency or share class disagree |

With three sources, the tool also points at the **probable faulty source**: if both counterparties agree with each other but not with Internal, Internal is suspect; if only one disagrees, that counterparty is suspect.

### Run

```bash
python nav_break_computation.py                               # real data in data/
python nav_break_computation.py --data-dir samples/breaks_demo  # demo with breaks
python nav_break_computation.py --warn-bps 2 --break-bps 10     # custom tolerances
```

Output: `output/step2_breaks_YYYY-MM-DD.xlsx` with three sheets — **Summary** (one row per fund), **Break Detail** (one row per fund × comparison) and **Exceptions** (warnings and breaks with action notes).

### Results on the sample date (30/05/2026)

- `data/` — 15 funds reconciled, **15 MATCHED** (all differences < 1 bp), 3 excluded by Step 1.
- `samples/breaks_demo/` — 10 MATCHED, 2 WARNING, 3 BREAK, each with the expected cause:

| ISIN | Injected anomaly | Detected |
|---|---|---|
| LU0003456789 | Stale price in Internal (−8 bps) | BREAK · PRICING · suspect INTERNAL |
| LU0010123456 | Administrator NAV +3 bps | WARNING · PRICING · suspect ADMIN |
| LU0006789012 | Custodian +1.5 bp on money market fund | WARNING (tight override) · suspect CUSTODIAN |
| LU0013456789 | 5 000 units subscribed, not booked internally | BREAK · CAPITAL_ACTIVITY · suspect CUSTODIAN |
| LU0012345678 | Custodian valued on 29/05 | BREAK · NAV_DATE · suspect CUSTODIAN |

### Tests

```bash
pytest -q        # 14 tests: thresholds, bps maths, triangulation, overrides, end-to-end demo
```

---

## Dependencies

- Python 3.10+
- `pandas`
- `openpyxl`
- `pytest` (tests only)

```bash
pip install -r requirements.txt
```

---

## Author

Youssef Ait Ali
