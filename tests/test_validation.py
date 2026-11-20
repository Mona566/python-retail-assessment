"""Tests for src/validation.py (Excel/workbook paths mocked)."""

from __future__ import annotations

from dataclasses import replace
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from src.analysis import AnalysisBundle
from src.excel_specs import PIVOT_SPECS
from src.validation import (
    REQUIRED_PNGS,
    REQUIRED_PIVOT_SHEETS,
    REQUIRED_SHEETS,
    _check_monthly_year_month_labels,
    _pivot_last_row_is_grand_total,
    run_validations,
    validate_analyses,
    validate_pngs,
    validate_prep,
    validate_workbook,
)


def test_required_pngs_and_sheets_follow_pivot_specs():
    assert REQUIRED_PNGS == [f"{spec['sheet']}.png" for spec in PIVOT_SPECS]
    assert REQUIRED_PIVOT_SHEETS == {spec["sheet"] for spec in PIVOT_SPECS}
    assert "Returns_Source" in REQUIRED_SHEETS
    assert REQUIRED_PIVOT_SHEETS <= REQUIRED_SHEETS


def test_validate_prep_passes_on_synthetic(clean_transactions, raw_transactions):
    from src.data_processing import apply_cross_sheet_dedupe

    kept, removed, _dedupe = apply_cross_sheet_dedupe(raw_transactions)
    audit = {
        "raw_rows": len(raw_transactions),
        "clean_rows": len(kept),
        "removed_cross_sheet_clones": len(removed),
    }
    # validate_prep expects enriched clean with same length as audit clean_rows
    from src.data_processing import enrich

    clean = enrich(kept)
    assert validate_prep(clean, audit) == []


def test_validate_analyses_passes(clean_transactions, analysis_bundle):
    errors = validate_analyses(clean_transactions, analysis_bundle)
    assert errors == [], errors


def test_validate_analyses_catches_uk_in_country(clean_transactions, analysis_bundle):
    bad = analysis_bundle
    # Inject UK into a copy of the country frame
    import pandas as pd
    from dataclasses import replace
    from src.analysis import AnalysisBundle

    country = bad.country.copy()
    country = pd.concat(
        [
            country,
            pd.DataFrame(
                [
                    {
                        "Country": "United Kingdom",
                        "gross_revenue": 1.0,
                        "distinct_invoices": 1,
                        "customers": 1,
                        "aov": 1.0,
                        "rank": 99,
                    }
                ]
            ),
        ],
        ignore_index=True,
    )
    tampered = AnalysisBundle(
        summary=bad.summary,
        product=bad.product,
        country=country,
        monthly=bad.monthly,
        returns=bad.returns,
        customer=bad.customer,
        timing=bad.timing,
        anomalies=bad.anomalies,
        narrative=bad.narrative,
    )
    errors = validate_analyses(clean_transactions, tampered)
    assert any("United Kingdom" in e for e in errors)


def test_validate_pngs_reports_missing(monkeypatch, tmp_path):
    from src import validation as v

    monkeypatch.setattr(v, "PNG_DIR", tmp_path)
    errors = validate_pngs()
    assert len(errors) == 6
    assert all("missing PNG" in e for e in errors)


def test_validate_prep_error_branches(clean_transactions):
    audit = {
        "raw_rows": 1,
        "clean_rows": len(clean_transactions) + 5,
        "removed_cross_sheet_clones": 0,
        "dedupe": {"matches_multiset_max": False},
    }
    errors = validate_prep(clean_transactions, audit)
    assert any("clean_rows mismatch" in e for e in errors)
    assert any("raw != clean" in e for e in errors)
    assert any("multiset max" in e for e in errors)

    # Missing columns
    slim = clean_transactions.drop(columns=["gross_revenue"])
    audit2 = {
        "raw_rows": len(slim),
        "clean_rows": len(slim),
        "removed_cross_sheet_clones": 0,
    }
    errors2 = validate_prep(slim, audit2)
    assert any("missing columns" in e for e in errors2)

    # Excel row limit
    from src import validation as v
    from src import data_processing as dp

    audit3 = {
        "raw_rows": dp.EXCEL_MAX_DATA_ROWS + 2,
        "clean_rows": dp.EXCEL_MAX_DATA_ROWS + 1,
        "removed_cross_sheet_clones": 1,
    }
    # Don't actually create huge frame — monkeypatch len via a tiny frame and override check
    # by temporarily lowering EXCEL_MAX_DATA_ROWS
    with patch.object(v, "EXCEL_MAX_DATA_ROWS", 5):
        # Need validate_prep to see patched limit — it imports from data_processing at module level
        pass
    with patch("src.validation.EXCEL_MAX_DATA_ROWS", 5):
        errors3 = validate_prep(clean_transactions, {
            "raw_rows": len(clean_transactions),
            "clean_rows": len(clean_transactions),
            "removed_cross_sheet_clones": 0,
        })
    assert any("Excel limit" in e for e in errors3)

    # Return-line identity / overlap
    bad = clean_transactions.copy()
    bad.loc[bad.index[0], "is_return_line"] = True
    bad.loc[bad.index[0], "is_confirmed_return"] = False
    bad.loc[bad.index[0], "is_stock_adjustment"] = False
    audit_ok = {
        "raw_rows": len(bad),
        "clean_rows": len(bad),
        "removed_cross_sheet_clones": 0,
    }
    assert any("is_return_line must equal" in e for e in validate_prep(bad, audit_ok))

    overlap = clean_transactions.copy()
    overlap.loc[overlap.index[0], "is_return_line"] = True
    overlap.loc[overlap.index[0], "is_confirmed_return"] = True
    overlap.loc[overlap.index[0], "is_stock_adjustment"] = True
    assert any("overlap" in e for e in validate_prep(overlap, audit_ok))


def test_validate_analyses_error_branches(clean_transactions, analysis_bundle):
    bad_summary = dict(analysis_bundle.summary)
    bad_summary["gross_revenue"] = -999.0
    bad_summary["distinct_invoices"] = -1
    bad_summary["net_revenue_after_cancellations"] = 0.0
    bad_summary["anomaly_thresholds"] = {"qty_hi": 1}  # missing keys
    bad_summary["peak_weekday"] = "wrong"
    bad_summary["peak_hour"] = -99

    huge_anom = analysis_bundle.anomalies.copy()
    if len(huge_anom) == 0:
        huge_anom = pd.DataFrame(
            [{"Invoice": "1", "Customer ID": 1, "reason": "x", "relevant_value": 1, "anomaly_score": 1}]
        )
    # pad to >20
    while len(huge_anom) <= 20:
        huge_anom = pd.concat([huge_anom, huge_anom], ignore_index=True)
    huge_anom = huge_anom.head(21)
    # drop a required col on a non-empty frame for missing-col branch
    anom_missing = huge_anom.drop(columns=["reason"])

    empty_product = analysis_bundle.product.iloc[0:0]
    empty_returns = analysis_bundle.returns.iloc[0:0]
    returns_unsorted = analysis_bundle.returns.copy()
    if len(returns_unsorted) >= 2:
        returns_unsorted = returns_unsorted.iloc[::-1].reset_index(drop=True)

    country = analysis_bundle.country.copy()
    if "missing_customer_id_revenue_share" in country.columns:
        country = country.drop(columns=["missing_customer_id_revenue_share"])

    tampered = AnalysisBundle(
        summary=bad_summary,
        product=empty_product,
        country=country,
        monthly=analysis_bundle.monthly,
        returns=empty_returns,
        customer=analysis_bundle.customer,
        timing=analysis_bundle.timing,
        anomalies=anom_missing,
        narrative=analysis_bundle.narrative,
    )
    errors = validate_analyses(clean_transactions, tampered)
    assert any("gross_revenue mismatch" in e for e in errors)
    assert any("distinct_invoices mismatch" in e for e in errors)
    assert any("net_revenue_after_cancellations" in e for e in errors)
    assert any("anomalies exceed" in e for e in errors)
    assert any("anomalies missing" in e for e in errors)
    assert any("anomaly_thresholds missing" in e for e in errors)
    assert any("product analysis empty" in e for e in errors)
    assert any("returns analysis empty" in e for e in errors)
    assert any("missing_customer_id_revenue_share" in e for e in errors)
    assert any("peak_weekday mismatch" in e for e in errors)
    assert any("peak_hour mismatch" in e for e in errors)

    # returns missing cols + unsorted
    returns_bad = analysis_bundle.returns.copy()
    if "return_line_rate" in returns_bad.columns:
        returns_bad = returns_bad.drop(columns=["return_line_rate"])
    tampered2 = replace(analysis_bundle, returns=returns_bad)
    errors2 = validate_analyses(clean_transactions, tampered2)
    assert any("returns analysis missing" in e for e in errors2)

    if len(returns_unsorted) >= 2 and "return_line_rate" in returns_unsorted.columns:
        tampered3 = replace(analysis_bundle, returns=returns_unsorted)
        errors3 = validate_analyses(clean_transactions, tampered3)
        assert any("not sorted" in e for e in errors3)

    # product helper mismatch
    bad_prod = analysis_bundle.product.copy()
    if len(bad_prod):
        clean2 = clean_transactions.copy()
        clean2["inv_flag_by_product"] = 0
        tampered4 = analysis_bundle
        errors4 = validate_analyses(clean2, tampered4)
        assert any("product helper mismatch" in e for e in errors4)


def test_validate_pngs_valid_blank_invalid(monkeypatch, tmp_path):
    from PIL import Image
    from src import validation as v
    from src.validation import REQUIRED_PNGS

    monkeypatch.setattr(v, "PNG_DIR", tmp_path)
    # Valid non-blank (gradient so extrema differ)
    for name in REQUIRED_PNGS:
        img = Image.new("RGB", (40, 30), color=(10, 20, 30))
        img.putpixel((0, 0), (255, 0, 0))
        img.save(tmp_path / name)
    assert validate_pngs() == []

    # Blank image (uniform color)
    Image.new("RGB", (40, 30), color=(0, 0, 0)).save(tmp_path / REQUIRED_PNGS[0])
    blank_errs = validate_pngs()
    assert any("blank" in e for e in blank_errs)

    # Invalid file
    (tmp_path / REQUIRED_PNGS[0]).write_bytes(b"not-a-png")
    inv_errs = validate_pngs()
    assert any("invalid screenshot" in e for e in inv_errs)

    # Wrong image format saved with .png extension
    Image.new("RGB", (40, 30), color=(1, 2, 3)).save(
        tmp_path / REQUIRED_PNGS[0], format="JPEG"
    )
    fmt_errs = validate_pngs()
    assert any("file is not a PNG" in e for e in fmt_errs)


def test_pivot_last_row_grand_total_edge_cases():
    empty = MagicMock()
    empty.range.value = None
    assert _pivot_last_row_is_grand_total(empty) is False

    bare = MagicMock()
    bare.range.value = [["Grand Total", 1], []]
    assert _pivot_last_row_is_grand_total(bare) is False

    scalar_last = MagicMock()
    scalar_last.range.value = [["a"], "not-a-list"]
    assert _pivot_last_row_is_grand_total(scalar_last) is False

    gt = MagicMock()
    gt.range.value = [["2011-03"], ["Grand Total", 9]]
    assert _pivot_last_row_is_grand_total(gt) is True


def test_monthly_year_month_label_skips_headers_and_sum_of():
    wb = MagicMock()
    pt = MagicMock()
    pt.range.value = [
        ["year_month"],
        ["Sum of gross_revenue"],
        ["Row Labels"],
        ["2011-03", 1],
    ]
    sheet = MagicMock()
    sheet.pivot_tables = [pt]
    wb.sheets = {"Pivot_Monthly_Performance": sheet}
    assert _check_monthly_year_month_labels(wb, {"Pivot_Monthly_Performance"}) == []

    pt.range.value = [["year_month"], ["Sum of gross_revenue"], ["Grand Total"]]
    errs = _check_monthly_year_month_labels(wb, {"Pivot_Monthly_Performance"})
    assert any("no year_month data rows" in e for e in errs)

    assert _check_monthly_year_month_labels(wb, set()) == []


def test_validate_workbook_mocked(tmp_path):
    missing = validate_workbook(tmp_path / "nope.xlsx")
    assert any("workbook missing" in e for e in missing)

    tiny = tmp_path / "tiny.xlsx"
    tiny.write_bytes(b"x" * 10)
    # Will try xlwings — mock it
    xw = MagicMock()

    class FakeApp:
        def __init__(self, *a, **k):
            self.display_alerts = True
            self.books = MagicMock()

        def quit(self):
            pass

    def open_partial(_path):
        wb = MagicMock()
        # Only some sheets
        sheet_names = ["Clean_Data", "Summary"]
        sheets = []
        for n in sheet_names:
            s = MagicMock()
            s.name = n
            sheets.append(s)
        wb.sheets = sheets
        # Also support wb.sheets[name]
        by_name = {s.name: s for s in sheets}
        type(wb).sheets = property(lambda self: _SheetCollection(by_name, sheets))
        return wb

    class _SheetCollection:
        def __init__(self, by_name, sheets):
            self._by_name = by_name
            self._sheets = sheets

        def __iter__(self):
            return iter(self._sheets)

        def __getitem__(self, key):
            return self._by_name[key]

    app = FakeApp()
    app.books.open.side_effect = open_partial
    xw.App = lambda *a, **k: app

    with patch("src.validation.xw", xw):
        errs = validate_workbook(tiny)
    assert any("unexpectedly small" in e for e in errs)
    assert any("missing sheets" in e for e in errs)

    # Full sheets, no pivots, grand total, bad year_month
    def open_full(_path):
        names = {
            "Clean_Data",
            "Summary",
            "Anomalies",
            "Returns_Source",
            "Pivot_Product_Revenue",
            "Pivot_Country_Performance",
            "Pivot_Monthly_Performance",
            "Pivot_Returns",
            "Pivot_Customer_Value",
            "Pivot_Sales_Timing",
        }
        sheets_list = []
        by_name = {}
        for n in names:
            s = MagicMock()
            s.name = n
            pt = MagicMock()
            if n == "Pivot_Sales_Timing":
                s.pivot_tables = []
            else:
                s.pivot_tables = [pt]
            if n == "Pivot_Monthly_Performance":
                pt.range.value = [
                    ["Year-month"],
                    ["not-a-date", 1],
                ]
            elif n.startswith("Pivot_") and n != "Pivot_Sales_Timing":
                pt.range.value = [["hdr"], ["Grand Total", 9]]
            by_name[n] = s
            sheets_list.append(s)
        wb = MagicMock()

        class Coll:
            def __iter__(self_inner):
                return iter(sheets_list)

            def __getitem__(self_inner, key):
                return by_name[key]

        wb.sheets = Coll()
        return wb

    big = tmp_path / "big.xlsx"
    big.write_bytes(b"x" * 2000)
    app.books.open.side_effect = open_full
    with patch("src.validation.xw", xw):
        errs2 = validate_workbook(big)
    assert any("no PivotTable" in e for e in errs2)
    assert any("Grand Total" in e for e in errs2)
    assert any("year_month label" in e for e in errs2)

    # Exception path
    app.books.open.side_effect = RuntimeError("com fail")
    with patch("src.validation.xw", xw):
        errs3 = validate_workbook(big)
    assert any("workbook validation error" in e for e in errs3)

    # Good monthly label + empty pivot values for _pivot_last_row helper
    def open_ok_month(_path):
        names = {
            "Clean_Data",
            "Summary",
            "Anomalies",
            "Returns_Source",
            "Pivot_Product_Revenue",
            "Pivot_Country_Performance",
            "Pivot_Monthly_Performance",
            "Pivot_Returns",
            "Pivot_Customer_Value",
            "Pivot_Sales_Timing",
        }
        sheets_list = []
        by_name = {}
        for n in names:
            s = MagicMock()
            s.name = n
            pt = MagicMock()
            s.pivot_tables = [pt]
            if n == "Pivot_Monthly_Performance":
                pt.range.value = [["Year-month"], ["2011-03", 10]]
            else:
                pt.range.value = [["hdr"], ["row", 1]]
            by_name[n] = s
            sheets_list.append(s)
        wb = MagicMock()

        class Coll:
            def __iter__(self_inner):
                return iter(sheets_list)

            def __getitem__(self_inner, key):
                return by_name[key]

        wb.sheets = Coll()
        return wb

    app.books.open.side_effect = open_ok_month
    with patch("src.validation.xw", xw):
        assert validate_workbook(big) == []


def test_run_validations(clean_transactions, analysis_bundle, monkeypatch, tmp_path):
    from src import validation as v

    audit = {
        "raw_rows": len(clean_transactions) + 1,
        "clean_rows": len(clean_transactions),
        "removed_cross_sheet_clones": 1,
        "dedupe": {"matches_multiset_max": True},
    }
    monkeypatch.setattr(v, "PNG_DIR", tmp_path)
    monkeypatch.setattr(v, "validate_workbook", lambda p: [])
    monkeypatch.setattr(v, "validate_pngs", lambda: [])
    assert run_validations(clean_transactions, audit, analysis_bundle, tmp_path / "wb.xlsx") == []

    errs = run_validations(
        clean_transactions,
        {"raw_rows": 0, "clean_rows": 999, "removed_cross_sheet_clones": 0},
        analysis_bundle,
        None,
        check_pngs=False,
    )
    assert errs

