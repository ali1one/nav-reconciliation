"""
nav_coverage_check.py
======================
Author  : Youssef Ait Ali
Purpose : Read NAV data from three independent sources (Internal system,
          Fund Administrator, Custodian), merge them on ISIN, and produce
          a coverage matrix that flags every fund missing from any source.

Usage
-----
    python nav_coverage_check.py

Expected input files (edit paths in CONFIG below if needed)
-----------------------------------------------------------
    data/internal/   internal_nav_YYYYMMDD.xlsx       (Excel)
    data/admin/      admin_nav_report_YYYYMMDD.csv    (CSV)
    data/custodian/  custodian_valuations_YYYYMMDD.csv (CSV)

Output
------
    output/step1_coverage_YYYY-MM-DD.xlsx
        Sheet "Coverage Matrix"  — one row per ISIN, ✓/✗ per source
        Sheet "Missing Funds"    — only the incomplete ISINs with action notes

Dependencies : pandas, openpyxl
"""

import os
import glob
import logging
import pandas as pd
from datetime import date

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

SOURCES = ["INTERNAL", "ADMIN", "CUSTODIAN"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — LOAD & NORMALISE EACH SOURCE
# ─────────────────────────────────────────────────────────────────────────────
#
# Each source uses different column names, date formats, and file types.
# Every loader reads its raw file and returns a DataFrame with the same
# standardised columns:
#
#   isin | fund_name | share_class | currency | nav_date |
#   total_nav | shares | navps | source | source_file
#
# ─────────────────────────────────────────────────────────────────────────────

DATE_FORMATS = [
    "%d/%m/%Y",    # internal:   30/05/2026
    "%Y-%m-%d",    # admin:      2026-05-30
    "%d-%b-%Y",    # custodian:  30-May-2026
    "%d/%m/%y",
    "%m/%d/%Y",
]

def _parse_date(val) -> date:
    """Try every known format until one works."""
    s = str(val).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.to_datetime(s, format=fmt).date()
        except Exception:
            continue
    raise ValueError(f"Cannot parse date: {s!r}")


def _latest_file(folder: str, pattern: str) -> str:
    """Return the most recent file matching a glob pattern."""
    files = sorted(glob.glob(os.path.join(folder, pattern)))
    if not files:
        raise FileNotFoundError(f"No files found: {os.path.join(folder, pattern)}")
    return files[-1]


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",", ""), errors="coerce")


# ── Internal system — Excel ──────────────────────────────────────────────────
# Columns: FundCode, ISIN, FundName, ShareClass, Ccy, ValuationDate,
#          TotalNAV, OutstandingShares, NAVPerShare, Domicile, Status

def load_internal() -> pd.DataFrame:
    path = _latest_file(os.path.join(DATA_DIR, "internal"), "*.xlsx")
    raw  = pd.read_excel(path, dtype=str)
    log.info(f"[INTERNAL]  {len(raw):>3} rows  ←  {os.path.basename(path)}")

    df = pd.DataFrame({
        "isin"        : raw["ISIN"],
        "fund_name"   : raw["FundName"],
        "share_class" : raw["ShareClass"],
        "currency"    : raw["Ccy"],
        "nav_date"    : raw["ValuationDate"].apply(_parse_date),
        "total_nav"   : _to_numeric(raw["TotalNAV"]),
        "shares"      : _to_numeric(raw["OutstandingShares"]),
        "source"      : "INTERNAL",
        "source_file" : os.path.basename(path),
    })
    df["navps"] = df["total_nav"] / df["shares"]
    return df


# ── Fund Administrator — CSV ──────────────────────────────────────────────────
# Columns: report_date, isin, sub_fund_name, share_class, currency,
#          net_asset_value, units_outstanding, nav_per_unit, administrator_ref

def load_admin() -> pd.DataFrame:
    path = _latest_file(os.path.join(DATA_DIR, "admin"), "*.csv")
    raw  = pd.read_csv(path, dtype=str)
    log.info(f"[ADMIN]     {len(raw):>3} rows  ←  {os.path.basename(path)}")

    df = pd.DataFrame({
        "isin"        : raw["isin"],
        "fund_name"   : raw["sub_fund_name"],
        "share_class" : raw["share_class"],
        "currency"    : raw["currency"],
        "nav_date"    : raw["report_date"].apply(_parse_date),
        "total_nav"   : _to_numeric(raw["net_asset_value"]),
        "shares"      : _to_numeric(raw["units_outstanding"]),
        "source"      : "ADMIN",
        "source_file" : os.path.basename(path),
    })
    df["navps"] = df["total_nav"] / df["shares"]
    return df


# ── Custodian — CSV ───────────────────────────────────────────────────────────
# Columns: ValueDate, SubFundISIN, SubFundName, Class, CCY,
#          TotalNetAssets, NumberOfUnits, UnitNAV, CustodianCode

def load_custodian() -> pd.DataFrame:
    path = _latest_file(os.path.join(DATA_DIR, "custodian"), "*.csv")
    raw  = pd.read_csv(path, dtype=str)
    log.info(f"[CUSTODIAN] {len(raw):>3} rows  ←  {os.path.basename(path)}")

    df = pd.DataFrame({
        "isin"        : raw["SubFundISIN"],
        "fund_name"   : raw["SubFundName"],
        "share_class" : raw["Class"],
        "currency"    : raw["CCY"],
        "nav_date"    : raw["ValueDate"].apply(_parse_date),
        "total_nav"   : _to_numeric(raw["TotalNetAssets"]),
        "shares"      : _to_numeric(raw["NumberOfUnits"]),
        "source"      : "CUSTODIAN",
        "source_file" : os.path.basename(path),
    })
    df["navps"] = df["total_nav"] / df["shares"]
    return df


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — BUILD COVERAGE MATRIX
# ─────────────────────────────────────────────────────────────────────────────
#
# Outer-join all three sources on ISIN.
# Result: one row per ISIN, with a ✓/✗ flag per source.
#
# ─────────────────────────────────────────────────────────────────────────────

def build_coverage(
    df_int:  pd.DataFrame,
    df_adm:  pd.DataFrame,
    df_cust: pd.DataFrame,
) -> pd.DataFrame:

    # Keep only the identifying columns + NAV for the merge
    def _slim(df: pd.DataFrame, src: str) -> pd.DataFrame:
        return df[["isin", "fund_name", "share_class", "currency",
                   "total_nav", "navps", "source_file"]].copy()\
                  .rename(columns={
                      "total_nav"   : f"total_nav_{src}",
                      "navps"       : f"navps_{src}",
                      "source_file" : f"file_{src}",
                  })

    slim_int  = _slim(df_int,  "INTERNAL")
    slim_adm  = _slim(df_adm,  "ADMIN")
    slim_cust = _slim(df_cust, "CUSTODIAN")

    # Outer join: every ISIN seen in any source gets a row
    merged = slim_int\
        .merge(slim_adm,  on=["isin", "fund_name", "share_class", "currency"], how="outer")\
        .merge(slim_cust, on=["isin", "fund_name", "share_class", "currency"], how="outer")

    # Presence flags
    for src in SOURCES:
        merged[f"in_{src}"] = merged[f"total_nav_{src}"].notna()

    merged["sources_present"] = merged[[f"in_{s}" for s in SOURCES]].sum(axis=1).astype(int)
    merged["sources_missing"]  = 3 - merged["sources_present"]
    merged["missing_from"]     = merged.apply(
        lambda r: [s for s in SOURCES if not r[f"in_{s}"]], axis=1
    )

    return merged.sort_values("isin").reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — CLASSIFY MISSING FUNDS
# ─────────────────────────────────────────────────────────────────────────────
#
# Categories
# ----------
# LATE_FILE       fund present in ≥2 sources; absent from 1 counterparty feed
#                 → most common; chase counterparty for resubmission
# NOT_ONBOARDED   fund at admin/custodian but absent from internal system
#                 → new fund not yet set up internally
# NEW_FUND_SETUP  fund in internal but absent from admin and/or custodian
#                 → counterparty setup still in progress
# ORPHAN          fund in only 1 source → likely a data/setup error
#
# ─────────────────────────────────────────────────────────────────────────────

CATEGORY_ICON = {
    "LATE_FILE"     : "⏱",
    "NOT_ONBOARDED" : "🆕",
    "NEW_FUND_SETUP": "⚙",
    "ORPHAN"        : "👤",
}

ACTION = {
    "LATE_FILE"     : "Fund absent from {missing} feed. Chase counterparty for resubmission or confirm non-valuation day.",
    "NOT_ONBOARDED" : "Fund present at counterparties but absent from internal system. Verify onboarding status with Fund Operations.",
    "NEW_FUND_SETUP": "Fund in internal system but absent from {missing}. Confirm counterparty setup is in progress.",
    "ORPHAN"        : "Fund present in only 1 source ({present}). Likely a data error — investigate before reconciling.",
}

def _classify_row(row) -> tuple[str, str]:
    m = row["missing_from"]
    p = [s for s in SOURCES if s not in m]

    if row["sources_present"] == 1:
        cat = "ORPHAN"
        return cat, ACTION[cat].format(present=", ".join(p))

    if "INTERNAL" in m and len(m) == 1:
        cat = "NOT_ONBOARDED"
        return cat, ACTION[cat]

    if "INTERNAL" not in m and len(m) >= 1:
        cat = "LATE_FILE" if len(m) == 1 else "NEW_FUND_SETUP"
        return cat, ACTION[cat].format(missing=", ".join(m))

    cat = "LATE_FILE"
    return cat, ACTION[cat].format(missing=", ".join(m))


def classify_missing(coverage: pd.DataFrame) -> pd.DataFrame:
    incomplete = coverage[coverage["sources_present"] < 3].copy()
    if incomplete.empty:
        return incomplete

    incomplete[["category", "action"]] = incomplete.apply(
        _classify_row, axis=1, result_type="expand"
    )
    incomplete["missing_from_str"] = incomplete["missing_from"].apply(", ".join)

    return incomplete[[
        "isin", "fund_name", "share_class", "currency",
        "sources_present", "missing_from_str", "category", "action",
    ]].reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — PRINT REPORT
# ─────────────────────────────────────────────────────────────────────────────

def print_report(
    sources:  dict,
    coverage: pd.DataFrame,
    missing:  pd.DataFrame,
    nav_date: date,
) -> None:

    W = 98
    print(f"\n{'═'*W}")
    print(f"  FUND NAV RECONCILIATION — STEP 1: DATA INGESTION & COVERAGE CHECK")
    print(f"  NAV Date : {nav_date}  |  "
          f"Sources : Internal (xlsx) · Admin (csv) · Custodian (csv)")
    print(f"{'═'*W}\n")

    # ── Source summary ────────────────────────────────────────────────────────
    print("SOURCE FILES LOADED")
    print("─"*W)
    print(f"  {'Source':<14} {'File':<48} {'Funds':>6}  {'NAV Date':>12}")
    print("  " + "─"*90)
    for src, df in sources.items():
        fname = df["source_file"].iloc[0]
        ndate = str(df["nav_date"].iloc[0])
        print(f"  {src:<14} {fname:<48} {len(df):>6}  {ndate:>12}")

    # ── Coverage matrix ───────────────────────────────────────────────────────
    total_isins = len(coverage)
    print(f"\n\nCOVERAGE MATRIX — {total_isins} ISINs across all sources")
    print("─"*W)
    hdr = (
        f"  {'ISIN':<14} {'Fund Name':<32} {'Class':<7} {'CCY':<5}"
        f" {'INT':^5} {'ADM':^5} {'CUST':^5}"
        f" {'Present':>8} {'Missing':>8}"
    )
    print(hdr)
    print("  " + "─"*90)

    for _, r in coverage.iterrows():
        flag     = "  ◄" if r["sources_missing"] > 0 else ""
        int_m  = "✓" if r["in_INTERNAL"]  else "✗"
        adm_m  = "✓" if r["in_ADMIN"]     else "✗"
        cust_m = "✓" if r["in_CUSTODIAN"] else "✗"
        print(
            f"  {r['isin']:<14} {str(r['fund_name']):<32} "
            f"{str(r['share_class']):<7} {str(r['currency']):<5}"
            f" {int_m:^5} {adm_m:^5} {cust_m:^5}"
            f" {int(r['sources_present']):>8} {int(r['sources_missing']):>8}{flag}"
        )

    # ── Missing fund detail ───────────────────────────────────────────────────
    print(f"\n\nMISSING FUND ANALYSIS")
    print("─"*W)

    if missing.empty:
        print("  ✓  All funds present in all three sources. No issues.\n")
    else:
        print(f"  {len(missing)} ISIN(s) require attention before proceeding to break computation.\n")
        for cat in ["ORPHAN", "NOT_ONBOARDED", "NEW_FUND_SETUP", "LATE_FILE"]:
            sub = missing[missing["category"] == cat]
            if sub.empty:
                continue
            icon = CATEGORY_ICON[cat]
            print(f"  {icon}  {cat}  ({len(sub)} fund{'s' if len(sub) > 1 else ''})")
            print("  " + "─"*85)
            for _, r in sub.iterrows():
                print(f"    ISIN         : {r['isin']}")
                print(f"    Fund         : {r['fund_name']}  [{r['share_class']}]  {r['currency']}")
                print(f"    Missing from : {r['missing_from_str']}")
                print(f"    Action       : {r['action']}")
                print()

    # ── Summary ───────────────────────────────────────────────────────────────
    complete   = (coverage["sources_present"] == 3).sum()
    incomplete = total_isins - complete

    print("SUMMARY")
    print("─"*W)
    print(f"  Total ISINs in universe  : {total_isins}")
    print(f"  ✓  Complete  (3/3 sources): {complete}")
    print(f"  ◄  Incomplete (<3 sources): {incomplete}")
    if not missing.empty:
        for cat, icon in CATEGORY_ICON.items():
            n = (missing["category"] == cat).sum()
            if n:
                print(f"       {icon}  {cat:<20}: {n}")
    print(f"\n  Funds eligible for break computation : {complete}")
    print(f"\n{'═'*W}")
    print(f"  ✓  Step 1 complete. Review missing funds above, then proceed to Step 2.")
    print(f"{'═'*W}\n")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — EXPORT TO EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def export_excel(coverage: pd.DataFrame, missing: pd.DataFrame, nav_date: date) -> str:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # ── Colour palette ────────────────────────────────────────────────────────
    C = {
        "header" : PatternFill("solid", fgColor="1F3864"),  # navy
        "ok"     : PatternFill("solid", fgColor="C6EFCE"),  # light green
        "miss"   : PatternFill("solid", fgColor="FFC7CE"),  # light red
        "warn"   : PatternFill("solid", fgColor="FFEB9C"),  # light amber
        "late"   : PatternFill("solid", fgColor="E2EFDA"),  # pale green
        "new"    : PatternFill("solid", fgColor="BDD7EE"),  # light blue
        "orphan" : PatternFill("solid", fgColor="FCE4D6"),  # light peach
    }
    CAT_FILL = {
        "LATE_FILE"     : C["late"],
        "NOT_ONBOARDED" : C["new"],
        "NEW_FUND_SETUP": C["warn"],
        "ORPHAN"        : C["orphan"],
    }
    FH  = Font(bold=True, color="FFFFFF", size=9)   # white on navy header
    FD  = Font(color="000000", size=9)               # black on light cells
    AC  = Alignment(horizontal="center", vertical="center")
    AR  = Alignment(horizontal="right",  vertical="center")
    AL  = Alignment(horizontal="left",   vertical="center")
    thin   = Side(style="thin", color="D9D9D9")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Build display frames ──────────────────────────────────────────────────
    cov_display = pd.DataFrame({
        "ISIN"           : coverage["isin"],
        "Fund Name"      : coverage["fund_name"],
        "Share Class"    : coverage["share_class"],
        "Currency"       : coverage["currency"],
        "Internal"       : coverage["in_INTERNAL"].map({True: "✓", False: "✗"}),
        "Admin"          : coverage["in_ADMIN"].map({True: "✓", False: "✗"}),
        "Custodian"      : coverage["in_CUSTODIAN"].map({True: "✓", False: "✗"}),
        "Sources Present": coverage["sources_present"],
        "Sources Missing": coverage["sources_missing"],
        "Status"         : coverage["sources_missing"].map(
            lambda x: "COMPLETE" if x == 0 else "INCOMPLETE"
        ),
    })

    path = os.path.join(OUTPUT_DIR, f"step1_coverage_{nav_date}.xlsx")

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        cov_display.to_excel(writer, sheet_name="Coverage Matrix", index=False)
        if not missing.empty:
            missing.to_excel(writer, sheet_name="Missing Funds", index=False)

    # ── Formatting pass ───────────────────────────────────────────────────────
    wb = load_workbook(path)

    # Coverage Matrix sheet
    ws = wb["Coverage Matrix"]
    int_col  = cov_display.columns.get_loc("Internal")  + 1
    adm_col  = cov_display.columns.get_loc("Admin")     + 1
    cust_col = cov_display.columns.get_loc("Custodian") + 1
    stat_col = cov_display.columns.get_loc("Status")    + 1

    for cell in ws[1]:
        cell.fill = C["header"]; cell.font = FH
        cell.border = BORDER;    cell.alignment = AC

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = FD; cell.border = BORDER; cell.alignment = AL
        # Colour source flag cells
        for col_idx, fill_ok, fill_bad in [
            (int_col,  C["ok"], C["miss"]),
            (adm_col,  C["ok"], C["miss"]),
            (cust_col, C["ok"], C["miss"]),
        ]:
            cell = ws.cell(row=row[0].row, column=col_idx)
            cell.fill = fill_ok if cell.value == "✓" else fill_bad
            cell.alignment = AC
        # Status cell
        s_cell = ws.cell(row=row[0].row, column=stat_col)
        s_cell.fill = C["ok"] if s_cell.value == "COMPLETE" else C["miss"]
        s_cell.alignment = AC

    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max(len(str(c.value or "")) for c in col) + 3, 38)

    # Missing Funds sheet
    if "Missing Funds" in wb.sheetnames:
        ws2 = wb["Missing Funds"]
        cat_col = list(missing.columns).index("category") + 1

        for cell in ws2[1]:
            cell.fill = C["header"]; cell.font = FH
            cell.border = BORDER;    cell.alignment = AC

        for row in ws2.iter_rows(min_row=2):
            cat_val = ws2.cell(row=row[0].row, column=cat_col).value
            row_fill = CAT_FILL.get(cat_val, C["warn"])
            for cell in row:
                cell.font = FD; cell.border = BORDER
                cell.fill = row_fill
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        for col in ws2.columns:
            ws2.column_dimensions[get_column_letter(col[0].column)].width = \
                min(max(len(str(c.value or "")) for c in col) + 3, 55)

    wb.save(path)
    return path


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":

    # 1. Load all three sources
    log.info("Loading source files...")
    df_int  = load_internal()
    df_adm  = load_admin()
    df_cust = load_custodian()

    sources  = {"INTERNAL": df_int, "ADMIN": df_adm, "CUSTODIAN": df_cust}
    nav_date = df_int["nav_date"].iloc[0]   # reference date from internal

    # 2. Build coverage matrix
    log.info("Building coverage matrix...")
    coverage = build_coverage(df_int, df_adm, df_cust)

    # 3. Classify missing funds
    log.info("Classifying missing funds...")
    missing  = classify_missing(coverage)

    # 4. Print report to console
    print_report(sources, coverage, missing, nav_date)

    # 5. Export to Excel
    log.info("Exporting to Excel...")
    out_path = export_excel(coverage, missing, nav_date)
    log.info(f"Output → {out_path}")
