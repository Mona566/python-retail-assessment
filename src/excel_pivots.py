"""Phase B: create, refresh, and style native Excel PivotTables via xlwings."""

from __future__ import annotations

from pathlib import Path

from src.analysis import AnalysisBundle
from src.constants import BusinessRules, PivotLayoutConst
from src.excel_specs import PIVOT_SPECS, PIVOT_VALUE_STYLE
from src.excel_write import _publication_sources
from src.logging_utils import get_logger

logger = get_logger()


def _style_pivot_fields(pt, spec: dict) -> None:
    """Apply number formats from PIVOT_VALUE_STYLE and widen header columns for PNG export."""
    for field, source in zip(pt.values, spec["values"], strict=True):
        _, number_format = PIVOT_VALUE_STYLE[source]
        field.number_format = number_format

    sheet = pt.parent
    row = PivotLayoutConst.HEADER_ROW_INDEX
    if spec.get("columns"):
        first = pt.range.column
        last = pt.range.last_cell.column
        sheet.range((row, first + 1), (row, last)).column_width = (
            PivotLayoutConst.COL_WIDTH_VALUE
        )
        sheet.range((row, first)).column_width = PivotLayoutConst.COL_WIDTH_ROW_LABEL
        return

    values = pt.range.options(ndim=2).value
    for offset, header in enumerate(values[1] if len(spec["values"]) > 1 else values[0]):
        text_width = len(str(header or "")) + PivotLayoutConst.HEADER_WIDTH_PAD
        if offset == 0 and spec["rows"] == "Product":
            text_width = (
                max(len(str(row[0] or "")) for row in values)
                + PivotLayoutConst.PRODUCT_WIDTH_PAD
            )
        sheet.range((row, pt.range.column + offset)).column_width = max(
            PivotLayoutConst.COL_WIDTH_MIN, text_width
        )


def _configure_grand_totals(pt, *, hide_bottom: bool) -> None:
    """
    On Excel PivotTables with row fields + value columns:
    - Bottom 'Grand Total' row  == ColumnGrand / show_column_grand_totals
    - Right-hand totals         == RowGrand / show_row_grand_totals

    Tries xlwings properties first, then the appscript API fallback used on
    Excel for Mac. When ``hide_bottom`` is True, raises if neither path can
    disable the bottom Grand Total (do not silently leave rates/helpers summed).
    """
    want_column_grand = not hide_bottom
    errors: list[str] = []
    column_via_xlwings = False
    appscript_ok = False

    try:
        pt.show_column_grand_totals = want_column_grand
        column_via_xlwings = True
    except Exception as exc:  # noqa: BLE001 — COM/appscript surface varies
        errors.append(f"xlwings show_column_grand_totals: {exc}")

    try:
        pt.show_row_grand_totals = False
    except Exception as exc:  # noqa: BLE001
        errors.append(f"xlwings show_row_grand_totals: {exc}")

    try:
        api = pt.api
        api.column_grand.set(want_column_grand)
        api.row_grand.set(False)
        appscript_ok = True
    except Exception as exc:  # noqa: BLE001
        errors.append(f"appscript column/row_grand: {exc}")

    if hide_bottom and not column_via_xlwings and not appscript_ok:
        raise RuntimeError(
            "Failed to hide bottom Grand Total "
            f"(xlwings and appscript both failed): {'; '.join(errors)}"
        )
    if errors:
        logger.debug("Grand total configure partial failures: %s", "; ".join(errors))


def _pivot_has_grand_total_row(pt) -> bool:
    """True if the last Pivot row label is Grand Total (should be hidden on rate pivots)."""
    values = pt.range.value
    if not values:
        return False
    last = values[-1]
    if not isinstance(last, list) or not last:
        return False
    return str(last[0]).strip().lower() == BusinessRules.PIVOT_GRAND_TOTAL_LABEL


def _create_native_pivots(wb, bundle: AnalysisBundle) -> list[str]:
    """Create native PivotTables on the open workbook; return created sheet names."""
    sources = {
        name: wb.sheets[name].used_range for name in _publication_sources(bundle)
    }

    created: list[str] = []
    for spec in PIVOT_SPECS:
        ws = wb.sheets[spec["sheet"]]
        ws.clear_contents()
        ws.range("A1").value = spec["sheet"].replace("_", " ")
        ws.range("A1").font.bold = True
        notes = []
        if spec.get("hide_bottom_grand_total"):
            notes.append(
                "Bottom Grand Total hidden — some values on this sheet "
                "are not additive across groups (helpers and/or rates)."
            )
        if spec.get("note"):
            notes.append(str(spec["note"]))
        if notes:
            ws.range("A2").value = " | ".join(notes)

        src = sources[spec["source_sheet"]]
        row_fields = spec["rows"]
        pt = ws.pivot_tables.add(
            source=src,
            destination=ws.range("A3"),
            name=spec["name"],
            rows=row_fields if isinstance(row_fields, str) else row_fields[0],
            values=spec["values"],
            layout="tabular",
        )
        if isinstance(row_fields, list) and len(row_fields) > 1:
            for extra in row_fields[1:]:
                try:
                    pt.rows.add(extra)
                except Exception as exc:  # noqa: BLE001 — Excel COM errors are opaque
                    # Re-raise with sheet/field context; preserve cause chain.
                    raise RuntimeError(
                        f"Could not add row field {extra!r} on {spec['sheet']}"
                    ) from exc
        if spec.get("columns"):
            pt.columns.add(spec["columns"])
        _configure_grand_totals(
            pt, hide_bottom=bool(spec.get("hide_bottom_grand_total"))
        )
        created.append(spec["sheet"])
        logger.info("Created pivot | name=%s sheet=%s", spec["name"], spec["sheet"])
    return created


def _refresh_and_style_pivots(app, output_path: Path, created: list[str]) -> None:
    """Re-open workbook, refresh/style pivots, save, and quit the Excel app.

    Failures (including persistent Grand Total rows on non-additive pivots)
    propagate — they must not be logged as warnings and treated as success.
    """
    wb2 = app.books.open(str(output_path))
    try:
        for name, spec in zip(created, PIVOT_SPECS, strict=True):
            pt = wb2.sheets[name].pivot_tables[0]
            pt.refresh()
            hide_bottom = bool(spec.get("hide_bottom_grand_total"))
            _configure_grand_totals(pt, hide_bottom=hide_bottom)
            if hide_bottom and _pivot_has_grand_total_row(pt):
                logger.warning(
                    "Grand Total still present after hide | sheet=%s action=retry",
                    name,
                )
                _configure_grand_totals(pt, hide_bottom=True)
                pt.refresh()
                if _pivot_has_grand_total_row(pt):
                    raise RuntimeError(
                        f"Grand Total row still present on {name} after hide "
                        "(non-additive rates/helpers must not show a bottom total)"
                    )
            _style_pivot_fields(pt, spec)

        wb2.save()
    finally:
        try:
            wb2.close()
        except Exception as exc:  # noqa: BLE001 — cleanup only
            logger.debug("Workbook close during refresh cleanup failed: %s", exc)
        try:
            app.quit()
        except Exception as exc:  # noqa: BLE001 — cleanup only
            logger.debug("Excel quit during refresh cleanup failed: %s", exc)
