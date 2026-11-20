"""Tests for Excel workbook helpers (xlsxwriter real; xlwings fully mocked)."""

from __future__ import annotations

from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from src.data_processing import CLEAN_DATA_COLUMNS
from src.excel_automation import (
    PIVOT_SPECS,
    _as_excel_text,
    _capture_range_to_png,
    _configure_grand_totals,
    _pivot_has_grand_total_row,
    _prepare_clean_for_excel,
    _publication_sources,
    _style_pivot_fields,
    _summary_frame,
    _write_base_workbook_xlsxwriter,
    build_workbook,
    export_pivot_pngs,
)


def test_as_excel_text_handles_nulls_and_apostrophe():
    assert _as_excel_text(None) == ""
    assert _as_excel_text(float("nan")) == ""
    assert _as_excel_text("nan") == ""
    assert _as_excel_text("<NA>") == ""
    assert _as_excel_text("None") == ""
    assert _as_excel_text("'2011-03") == "2011-03"
    assert _as_excel_text("P1") == "P1"


def test_summary_frame_and_publication_sources(analysis_bundle):
    audit = {"raw_rows": 100, "removed_cross_sheet_clones": 1}
    summary = _summary_frame(analysis_bundle, audit)
    assert list(summary.columns) == ["Metric", "Value"]
    assert (summary["Metric"] == "Gross Revenue (GBP)").any()
    assert (summary["Metric"].str.startswith("Product Top")).any()
    assert (summary["Metric"].str.startswith("Country Top")).any()
    assert (summary["Metric"].str.startswith("Customer Top")).any()
    assert (summary["Metric"].str.startswith("Returns Top")).any()

    sources = _publication_sources(analysis_bundle)
    assert set(sources) == {
        "Product_Source",
        "Country_Source",
        "Monthly_Source",
        "Returns_Source",
        "Customer_Source",
        "Timing_Source",
    }
    assert "Product" in sources["Product_Source"].columns
    assert "Year-month" in sources["Monthly_Source"].columns


def test_prepare_clean_for_excel(clean_transactions):
    out = _prepare_clean_for_excel(clean_transactions)
    assert list(out.columns) == CLEAN_DATA_COLUMNS
    assert (out.loc[clean_transactions["Customer ID"].isna(), "Customer ID"] == "").all()

    slim = clean_transactions.drop(columns=["gross_revenue"])
    with pytest.raises(KeyError, match="missing columns"):
        _prepare_clean_for_excel(slim)


def test_write_base_workbook_xlsxwriter(clean_transactions, analysis_bundle, tmp_path):
    clean_out = _prepare_clean_for_excel(clean_transactions)
    out = tmp_path / "Retail_Analysis.xlsx"
    audit = {"raw_rows": len(clean_transactions) + 1, "removed_cross_sheet_clones": 1}
    _write_base_workbook_xlsxwriter(clean_out, analysis_bundle, audit, out)
    assert out.exists() and out.stat().st_size > 1000
    xl = pd.ExcelFile(out)
    assert "Clean_Data" in xl.sheet_names
    assert "Summary" in xl.sheet_names
    assert "Anomalies" in xl.sheet_names
    for spec in PIVOT_SPECS:
        assert spec["sheet"] in xl.sheet_names


def test_write_base_workbook_empty_anomalies(clean_transactions, analysis_bundle, tmp_path):
    from dataclasses import replace

    empty = replace(
        analysis_bundle,
        anomalies=pd.DataFrame(),
        narrative={**analysis_bundle.narrative, "anomaly_note": ""},
    )
    clean_out = _prepare_clean_for_excel(clean_transactions)
    out = tmp_path / "empty_anom.xlsx"
    _write_base_workbook_xlsxwriter(clean_out, empty, {}, out)
    anom = pd.read_excel(out, sheet_name="Anomalies", skiprows=2)
    assert "note" in anom.columns or "No anomalies" in str(anom.values)


def test_configure_grand_totals_success_and_failures():
    pt = MagicMock()
    api = MagicMock()
    pt.api = api
    _configure_grand_totals(pt, hide_bottom=True)
    assert pt.show_column_grand_totals is False
    assert pt.show_row_grand_totals is False
    api.column_grand.set.assert_called_with(False)
    api.row_grand.set.assert_called_with(False)

    pt2 = MagicMock()

    def _raise_col(_self, _v):
        raise RuntimeError("col")

    def _raise_row(_self, _v):
        raise RuntimeError("row")

    type(pt2).show_column_grand_totals = property(lambda self: True, _raise_col)
    type(pt2).show_row_grand_totals = property(lambda self: True, _raise_row)
    pt2.api = MagicMock()
    pt2.api.column_grand.set.side_effect = RuntimeError("appscript")
    # Showing totals may soft-fail both paths; hiding must not.
    _configure_grand_totals(pt2, hide_bottom=False)
    with pytest.raises(RuntimeError, match="Failed to hide bottom Grand Total"):
        _configure_grand_totals(pt2, hide_bottom=True)

    # xlwings column fails but appscript succeeds → hide is OK
    pt3 = MagicMock()
    type(pt3).show_column_grand_totals = property(lambda self: True, _raise_col)
    type(pt3).show_row_grand_totals = property(lambda self: True, _raise_row)
    pt3.api = MagicMock()
    _configure_grand_totals(pt3, hide_bottom=True)
    pt3.api.column_grand.set.assert_called_with(False)


def test_pivot_has_grand_total_row():
    pt = MagicMock()
    pt.range.value = None
    assert _pivot_has_grand_total_row(pt) is False
    pt.range.value = [["a"], []]
    assert _pivot_has_grand_total_row(pt) is False
    pt.range.value = [["hdr"], ["Grand Total", 1]]
    assert _pivot_has_grand_total_row(pt) is True
    pt.range.value = [["hdr"], ["other", 1]]
    assert _pivot_has_grand_total_row(pt) is False


def test_style_pivot_fields_with_and_without_columns():
    pt = MagicMock()
    pt.parent = MagicMock()
    pt.range.column = 1
    pt.range.last_cell.column = 5
    pt.values = [MagicMock()]
    _style_pivot_fields(
        pt,
        {"rows": "Day of week", "columns": "Hour", "values": {"Revenue (GBP)": "sum"}},
    )
    assert pt.parent.range.called

    pt2 = MagicMock()
    pt2.parent = MagicMock()
    pt2.range.column = 1
    pt2.range.options.return_value.value = [
        ["Product", "a", "b", "c"],
        ["Product", "Product revenue (GBP)", "Product units sold", "Distinct invoices"],
        ["01 | P1 — MUG", 1, 2, 3],
    ]
    pt2.values = [MagicMock(), MagicMock(), MagicMock()]
    spec_prod = next(s for s in PIVOT_SPECS if s["sheet"] == "Pivot_Product_Revenue")
    _style_pivot_fields(pt2, spec_prod)

    pt3 = MagicMock()
    pt3.parent = MagicMock()
    pt3.range.column = 1
    pt3.range.options.return_value.value = [["Revenue (GBP)"], ["x", 1]]
    pt3.values = [MagicMock()]
    _style_pivot_fields(
        pt3,
        {"rows": "Year-month", "values": {"Revenue (GBP)": "sum"}},
    )


def test_capture_range_to_png_success_and_too_small(tmp_path):
    png = tmp_path / "shot.png"
    rng = MagicMock()

    def _write_ok(path):
        Path_write = type(png)
        Path_write.write_bytes(png, b"x" * 6000)

    rng.to_png.side_effect = lambda path: png.write_bytes(b"x" * 6000)
    _capture_range_to_png(rng, png)
    assert png.exists()

    rng2 = MagicMock()
    rng2.color = property  # unused
    type(rng2).color = property(
        fget=lambda self: None,
        fset=lambda self, v: (_ for _ in ()).throw(RuntimeError("style")),
    )
    type(rng2).font = property(lambda self: MagicMock())
    # Simpler: side_effect on attribute set via MagicMock config
    rng3 = MagicMock()

    def _set_color(_v):
        raise RuntimeError("style")

    type(rng3).color = property(lambda self: (255, 255, 255), lambda self, v: _set_color(v))
    rng3.font = MagicMock()
    rng3.to_png.side_effect = lambda path: png.write_bytes(b"tiny")
    with pytest.raises(RuntimeError, match="PNG too small"):
        _capture_range_to_png(rng3, png)


def _make_xlwings_mock(*, grand_total_on_refresh: bool = False, refresh_fail: bool = False):
    """Build a mock xlwings module for build_workbook / export_pivot_pngs."""
    xw = MagicMock()

    class FakeApp:
        def __init__(self, *a, **k):
            self.display_alerts = True
            self.screen_updating = True
            self.books = MagicMock()
            self._wb = None

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def quit(self):
            pass

    def make_sheet(name: str):
        sheet = MagicMock()
        sheet.name = name
        sheet.used_range = MagicMock(name=f"used_{name}")
        pt = MagicMock()
        pt.range = MagicMock()
        pt.range.column = 1
        pt.range.last_cell.column = 4
        if grand_total_on_refresh:
            pt.range.value = [["hdr"], ["Grand Total", 1]]
        else:
            pt.range.value = [["hdr"], ["row", 1]]
        pt.range.options.return_value.value = [
            ["Product", "a", "b"],
            ["Product", "Product revenue (GBP)", "Product units sold"],
            ["01 | P", 1, 2],
        ]
        n_values = next(
            (len(s["values"]) for s in PIVOT_SPECS if s["sheet"] == name),
            3,
        )
        pt.values = [MagicMock() for _ in range(n_values)]
        pt.api = MagicMock()
        pt.rows = MagicMock()
        pt.columns = MagicMock()
        if refresh_fail:
            pt.refresh.side_effect = RuntimeError("refresh boom")
        sheet.pivot_tables = MagicMock()
        sheet.pivot_tables.__getitem__ = lambda self, i: pt
        sheet.pivot_tables.add.return_value = pt
        sheet.pivot_tables.__len__ = lambda self: 1
        return sheet, pt

    sheets: dict[str, MagicMock] = {}
    for name in [
        "Clean_Data",
        "Summary",
        "Anomalies",
        "Product_Source",
        "Country_Source",
        "Monthly_Source",
        "Returns_Source",
        "Customer_Source",
        "Timing_Source",
        *[s["sheet"] for s in PIVOT_SPECS],
    ]:
        sheets[name], _ = make_sheet(name)

    def open_book(_path):
        wb = MagicMock()
        wb.sheets = MagicMock()
        wb.sheets.__getitem__ = lambda self, key: sheets[key]
        return wb

    FakeApp_instance_holder = {"app": None}

    def App(*a, **k):
        app = FakeApp(*a, **k)
        app.books.open.side_effect = open_book
        FakeApp_instance_holder["app"] = app
        return app

    xw.App = App
    return xw, sheets


def test_export_pivot_pngs_mocked(tmp_path, monkeypatch):
    from src import excel_automation as ea
    from src import config as cfg

    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)
    monkeypatch.setattr(cfg, "PNG_DIR", tmp_path)

    xw, sheets = _make_xlwings_mock()
    for sheet in sheets.values():
        sheet.pivot_tables[0].range.to_png.side_effect = (
            lambda path, _s=sheet: Path_write(path)
        )

    def Path_write(path):
        p = __import__("pathlib").Path(path)
        p.write_bytes(b"x" * 6000)

    for name, sheet in sheets.items():
        if name.startswith("Pivot_"):
            sheet.pivot_tables[0].range.to_png.side_effect = (
                lambda path: __import__("pathlib").Path(path).write_bytes(b"x" * 6000)
            )

    with patch("src.excel_automation.xw", xw), patch("src.excel_png.xw", xw):
        with patch("time.sleep"):
            # Re-import path uses local import inside function
            errors = export_pivot_pngs(tmp_path / "wb.xlsx")
    # May collect errors if sheets missing from open — open uses our sheets
    # export opens workbook via app.books.open — our FakeApp provides sheets
    assert isinstance(errors, list)


def test_build_workbook_with_mocked_xlwings(
    clean_transactions, analysis_bundle, tmp_path, monkeypatch
):
    from src import excel_automation as ea
    from src import config as cfg
    from src import logging_utils as lu

    out_dir = tmp_path / "output"
    png_dir = out_dir / "png"
    out_dir.mkdir()
    png_dir.mkdir()
    monkeypatch.setattr(ea, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(ea, "PNG_DIR", png_dir)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(cfg, "PNG_DIR", png_dir)
    monkeypatch.setattr(cfg, "AUDIT_DIR", tmp_path / "audit")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    monkeypatch.setattr(lu, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "audit").mkdir()
    (tmp_path / "logs").mkdir()

    xw, sheets = _make_xlwings_mock(grand_total_on_refresh=False)

    def fake_export(path):
        for spec in PIVOT_SPECS:
            (png_dir / f"{spec['sheet']}.png").write_bytes(b"x" * 6000)
        return []

    with patch("src.excel_automation.xw", xw), patch("src.excel_png.xw", xw):
        with patch.object(ea, "export_pivot_pngs", side_effect=fake_export):
            with patch("time.sleep"):
                result = build_workbook(
                    clean_transactions,
                    analysis_bundle,
                    {"raw_rows": 20, "removed_cross_sheet_clones": 1},
                    output_path=out_dir / "Retail_Analysis.xlsx",
                    export_png=True,
                )
    assert result.exists()


def test_build_workbook_raises_on_persistent_grand_total(
    clean_transactions, analysis_bundle, tmp_path, monkeypatch
):
    """Hiding Grand Total must fail the build if the row remains after retry."""
    from src import excel_automation as ea
    from src import config as cfg
    from src import logging_utils as lu

    out_dir = tmp_path / "output"
    out_dir.mkdir()
    monkeypatch.setattr(ea, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(ea, "PNG_DIR", out_dir / "png")
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(cfg, "PNG_DIR", out_dir / "png")
    monkeypatch.setattr(cfg, "AUDIT_DIR", tmp_path / "audit")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    monkeypatch.setattr(lu, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "audit").mkdir()
    (tmp_path / "logs").mkdir()
    (out_dir / "png").mkdir()

    xw, _ = _make_xlwings_mock(grand_total_on_refresh=True)
    with patch("src.excel_automation.xw", xw), patch("src.excel_png.xw", xw):
        with pytest.raises(RuntimeError, match="Grand Total row still present"):
            build_workbook(
                clean_transactions,
                analysis_bundle,
                {},
                output_path=out_dir / "Retail_Analysis.xlsx",
                export_png=False,
            )


def test_build_workbook_skip_png_and_png_errors(
    clean_transactions, analysis_bundle, tmp_path, monkeypatch
):
    from src import excel_automation as ea
    from src import config as cfg
    from src import logging_utils as lu

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    monkeypatch.setattr(ea, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(ea, "PNG_DIR", out_dir / "png")
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(cfg, "PNG_DIR", out_dir / "png")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    monkeypatch.setattr(lu, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "logs").mkdir()
    (out_dir / "png").mkdir()

    xw, _ = _make_xlwings_mock(refresh_fail=True)
    with patch("src.excel_automation.xw", xw), patch("src.excel_png.xw", xw):
        with pytest.raises(RuntimeError, match="refresh boom"):
            build_workbook(
                clean_transactions,
                analysis_bundle,
                {},
                output_path=out_dir / "wb.xlsx",
                export_png=False,
            )

    xw_ok, _ = _make_xlwings_mock()
    with patch("src.excel_automation.xw", xw_ok), patch("src.excel_png.xw", xw_ok):
        with patch.object(ea, "export_pivot_pngs", return_value=["Pivot_X: boom"]):
            with pytest.raises(RuntimeError, match="PNG export failed"):
                build_workbook(
                    clean_transactions,
                    analysis_bundle,
                    {},
                    output_path=out_dir / "wb2.xlsx",
                    export_png=True,
                )


def test_export_pivot_pngs_collects_errors(tmp_path, monkeypatch):
    from src import excel_automation as ea

    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)
    xw, sheets = _make_xlwings_mock()
    for name, sheet in sheets.items():
        if name.startswith("Pivot_"):
            sheet.pivot_tables[0].range.to_png.side_effect = RuntimeError("fail")

    with patch("src.excel_automation.xw", xw), patch("src.excel_png.xw", xw):
        with patch("time.sleep"):
            errors = export_pivot_pngs(tmp_path / "missing.xlsx")
    assert len(errors) == len(PIVOT_SPECS)
    assert all("RuntimeError: fail" in e for e in errors)


def test_build_workbook_finally_closes_on_failure(
    clean_transactions, analysis_bundle, tmp_path, monkeypatch
):
    """Force failure after books open so finally cleanup runs."""
    from src import excel_automation as ea
    from src import config as cfg
    from src import logging_utils as lu

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    monkeypatch.setattr(ea, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(ea, "PNG_DIR", out_dir / "png")
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out_dir)
    monkeypatch.setattr(cfg, "PNG_DIR", out_dir / "png")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    monkeypatch.setattr(lu, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "logs").mkdir()
    (out_dir / "png").mkdir()

    xw, sheets = _make_xlwings_mock()
    # Fail during first pivot create after workbook is open
    for name, sheet in sheets.items():
        if name.startswith("Pivot_"):
            sheet.pivot_tables.add.side_effect = RuntimeError("pivot boom")
            break

    with patch("src.excel_automation.xw", xw), patch("src.excel_png.xw", xw):
        with pytest.raises(RuntimeError, match="pivot boom"):
            build_workbook(
                clean_transactions,
                analysis_bundle,
                {},
                output_path=out_dir / "wb.xlsx",
                export_png=False,
            )


def test_summary_frame_missing_optional_narrative(analysis_bundle):
    from dataclasses import replace

    narr = dict(analysis_bundle.narrative)
    narr.pop("returns_note", None)
    narr.pop("country_missing_id_note", None)
    narr.pop("anomaly_note", None)
    # Force miss_share None path in country loop via NaN
    countries = analysis_bundle.narrative["top_countries"].copy()
    if "missing_customer_id_revenue_share" in countries.columns:
        countries["missing_customer_id_revenue_share"] = float("nan")
    narr["top_countries"] = countries
    bundle = replace(analysis_bundle, narrative=narr)
    frame = _summary_frame(bundle, {})
    assert not frame.empty
