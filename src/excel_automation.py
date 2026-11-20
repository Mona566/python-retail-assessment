"""Build Retail_Analysis.xlsx with Clean_Data, Summary, Anomalies, and native PivotTables."""

from __future__ import annotations

import time
from pathlib import Path

import pandas as pd
import xlwings as xw

from src.analysis import AnalysisBundle
from src.config import OUTPUT_DIR, OUTPUT_XLSX_NAME, PNG_DIR, ensure_dirs
from src.data_processing import validate_enrich_schema
from src.excel_pivots import (
    _configure_grand_totals,
    _create_native_pivots,
    _pivot_has_grand_total_row,
    _refresh_and_style_pivots,
    _style_pivot_fields,
)
from src.excel_png import _capture_range_to_png, export_pivot_pngs
from src.excel_specs import (
    CLEAN_DATA_EXCEL_RENAME,
    PIVOT_SPECS,
    PIVOT_VALUE_STYLE,
    RETURNS_SOURCE_EXCEL_RENAME,
)
from src.excel_summary import _summary_frame
from src.excel_write import (
    _apply_clean_column_formats,
    _as_excel_text,
    _prepare_clean_for_excel,
    _publication_sources,
    _write_base_workbook_xlsxwriter,
)
from src.logging_utils import get_logger, timed_phase

logger = get_logger()

__all__ = [
    "CLEAN_DATA_EXCEL_RENAME",
    "PIVOT_SPECS",
    "PIVOT_VALUE_STYLE",
    "RETURNS_SOURCE_EXCEL_RENAME",
    "_apply_clean_column_formats",
    "_as_excel_text",
    "_capture_range_to_png",
    "_configure_grand_totals",
    "_create_native_pivots",
    "_pivot_has_grand_total_row",
    "_prepare_clean_for_excel",
    "_publication_sources",
    "_refresh_and_style_pivots",
    "_style_pivot_fields",
    "_summary_frame",
    "_write_base_workbook_xlsxwriter",
    "build_workbook",
    "export_pivot_pngs",
]


def build_workbook(
    clean: pd.DataFrame,
    bundle: AnalysisBundle,
    audit: dict,
    output_path: Path | None = None,
    export_png: bool = True,
) -> Path:
    """
    Build Retail_Analysis.xlsx: Clean_Data + Summary + Anomalies + native PivotTables.

    Phase A writes tables via xlsxwriter; Phase B creates/refreshes pivots with xlwings.
    Optionally exports six Pivot PNGs (requires visible Excel). Returns the output path.
    """
    schema_errors = validate_enrich_schema(clean)
    if schema_errors:
        raise ValueError(f"build_workbook requires enriched ledger: {schema_errors[0]}")
    ensure_dirs()
    output_path = Path(output_path) if output_path else OUTPUT_DIR / OUTPUT_XLSX_NAME
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with timed_phase("excel_prepare_clean", rows=len(clean)) as meta:
        clean_out = _prepare_clean_for_excel(clean)
        meta["columns"] = clean_out.shape[1]

    t0 = time.perf_counter()
    with timed_phase("excel_write_base", rows=len(clean_out)):
        _write_base_workbook_xlsxwriter(clean_out, bundle, audit, output_path)

    app = None
    wb = None
    try:
        with timed_phase("excel_create_pivots") as pivot_meta:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(str(output_path))

            created = _create_native_pivots(wb, bundle)

            logger.info("Saving workbook with pivots | path=%s", output_path)
            wb.save()
            wb.close()
            wb = None
            pivot_meta["pivots_created"] = len(created)

        with timed_phase("excel_refresh_style", pivots=len(created)):
            _refresh_and_style_pivots(app, output_path, created)
            app = None

        if export_png:
            with timed_phase("excel_export_pngs"):
                png_errors = export_pivot_pngs(output_path)
                if png_errors:
                    raise RuntimeError(
                        "PNG export failed for one or more PivotTables (refusing to continue "
                        f"to email with missing/blank screenshots): {png_errors}"
                    )
        else:
            logger.info("PNG export skipped | export_png=False")

        logger.info(
            "Workbook complete | seconds=%.1f path=%s mb=%.1f",
            time.perf_counter() - t0,
            output_path,
            output_path.stat().st_size / 1e6,
        )
        return output_path
    finally:
        # Cleanup only — never mask the primary exception from the try body.
        if wb is not None:
            try:
                wb.close()
            except Exception as exc:  # noqa: BLE001
                logger.debug("Workbook close during build cleanup failed: %s", exc)
        if app is not None:
            try:
                app.quit()
            except Exception as exc:  # noqa: BLE001
                logger.debug("Excel quit during build cleanup failed: %s", exc)
