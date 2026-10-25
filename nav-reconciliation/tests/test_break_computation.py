"""Unit tests for Step 2 — run with:  pytest -q"""
import os
import sys
from datetime import date

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import nav_break_computation as s2  # noqa: E402

D = date(2026, 5, 30)


def _src(name, navps=10.0, shares=1_000_000, nav_date=D, ccy="EUR", isin="LU0000000001"):
    return pd.DataFrame([{
        "isin": isin, "fund_name": "Test Fund", "share_class": "A-EUR", "currency": ccy,
        "nav_date": nav_date, "total_nav": navps * shares, "shares": shares,
        "navps": navps, "source": name, "source_file": f"{name}.csv",
    }])


def _run(internal=None, admin=None, custodian=None, tol=s2.Tolerance()):
    sources = {"INTERNAL": _src("INTERNAL") if internal is None else internal,
               "ADMIN": _src("ADMIN") if admin is None else admin,
               "CUSTODIAN": _src("CUSTODIAN") if custodian is None else custodian}
    wide, excluded = s2.build_wide(sources)
    breaks = s2.compute_breaks(wide, tol)
    return s2.summarise(breaks, wide, tol).iloc[0], breaks


def test_identical_sources_match():
    row, breaks = _run()
    assert row["overall_status"] == "MATCHED"
    assert (breaks["navps_diff_bps"] == 0).all()
    assert row["suspect_source"] == ""


def test_bps_computation():
    _, breaks = _run(admin=_src("ADMIN", navps=10.003))      # +3 bps
    adm = breaks[breaks["counterparty"] == "ADMIN"].iloc[0]
    assert adm["navps_diff_bps"] == pytest.approx(3.0)


@pytest.mark.parametrize("navps, expected", [
    (10.0009, "MATCHED"),   # 0.9 bp
    (10.003, "WARNING"),    # 3 bp
    (10.006, "BREAK"),      # 6 bp
])
def test_navps_thresholds(navps, expected):
    row, _ = _run(admin=_src("ADMIN", navps=navps))
    assert row["status_ADMIN"] == expected


def test_one_counterparty_off_is_suspect():
    row, _ = _run(custodian=_src("CUSTODIAN", navps=10.01))
    assert row["suspect_source"] == "CUSTODIAN"
    assert row["root_cause"] == "PRICING"


def test_both_counterparties_agree_internal_is_suspect():
    row, _ = _run(admin=_src("ADMIN", navps=10.01), custodian=_src("CUSTODIAN", navps=10.01))
    assert row["overall_status"] == "BREAK"
    assert row["suspect_source"] == "INTERNAL"


def test_unit_difference_is_capital_activity():
    row, _ = _run(custodian=_src("CUSTODIAN", shares=1_005_000))
    assert row["overall_status"] == "BREAK"
    assert row["root_cause"] == "CAPITAL_ACTIVITY"


def test_nav_date_mismatch_breaks():
    row, _ = _run(admin=_src("ADMIN", nav_date=date(2026, 5, 29)))
    assert row["status_ADMIN"] == "BREAK"
    assert row["root_cause"] == "NAV_DATE"


def test_currency_mismatch_breaks():
    row, _ = _run(admin=_src("ADMIN", ccy="USD"))
    assert row["root_cause"] == "STATIC_DATA"


def test_override_tightens_tolerance(monkeypatch):
    monkeypatch.setattr(s2, "TOLERANCE_OVERRIDES",
                        {"LU0000000001": s2.Tolerance(0.5, 2.0, 1.0)})
    row, _ = _run(admin=_src("ADMIN", navps=10.003))          # 3 bp > 2 bp override
    assert row["status_ADMIN"] == "BREAK"


def test_incomplete_funds_are_excluded():
    extra = pd.concat([_src("INTERNAL"), _src("INTERNAL", isin="LU0000000002")])
    wide, excluded = s2.build_wide({"INTERNAL": extra, "ADMIN": _src("ADMIN"),
                                    "CUSTODIAN": _src("CUSTODIAN")})
    assert list(wide["isin"]) == ["LU0000000001"]
    assert excluded == ["LU0000000002"]


def test_duplicate_isin_raises():
    dup = pd.concat([_src("ADMIN"), _src("ADMIN")])
    with pytest.raises(ValueError, match="duplicated"):
        s2.build_wide({"INTERNAL": _src("INTERNAL"), "ADMIN": dup,
                       "CUSTODIAN": _src("CUSTODIAN")})


def test_demo_dataset_end_to_end(tmp_path):
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sources = s2.load_sources(os.path.join(root, "samples", "breaks_demo"))
    wide, _ = s2.build_wide(sources)
    summary = s2.summarise(s2.compute_breaks(wide), wide).set_index("isin")
    assert summary.loc["LU0003456789", "suspect_source"] == "INTERNAL"
    assert summary.loc["LU0010123456", "overall_status"] == "WARNING"
    assert summary.loc["LU0012345678", "root_cause"] == "NAV_DATE"
    assert summary.loc["LU0013456789", "root_cause"] == "CAPITAL_ACTIVITY"
    assert (summary["overall_status"] == "MATCHED").sum() == 10
