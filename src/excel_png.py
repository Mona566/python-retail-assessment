"""Export PivotTable ranges from Retail_Analysis.xlsx to PNG screenshots."""

from __future__ import annotations

import time
from pathlib import Path

import xlwings as xw

from src import config
from src.config import PNG_DIR  # noqa: F401 — re-export; runtime uses config.PNG_DIR for monkeypatch
from src.constants import ExcelPngConst
from src.excel_specs import PIVOT_SPECS
from src.logging_utils import get_logger

logger = get_logger()


def _capture_range_to_png(rng, png: Path) -> None:
    """Export an Excel range to PNG; refuse tiny/blank captures.

    Light-style forcing is best-effort (some ranges reject color writes); the
    subsequent ``to_png`` / size check still fail loudly on blank captures.
    """
    try:
        rng.color = ExcelPngConst.FORCE_FILL_RGB
        rng.font.color = ExcelPngConst.FORCE_FONT_RGB
    except Exception as style_exc:  # noqa: BLE001 — style is optional for capture
        logger.debug("Could not force light style before PNG: %s", style_exc)
    if png.exists():
        png.unlink()
    rng.to_png(str(png))
    size = png.stat().st_size if png.exists() else 0
    if size < ExcelPngConst.MIN_BYTES:
        raise RuntimeError(f"PNG too small ({size} bytes)")


def export_pivot_pngs(workbook_path: Path) -> list[str]:
    """
    Capture each PIVOT_SPECS sheet's PivotTable range to PNG_DIR/{sheet}.png.

    Uses a visible Excel session (required for reliable to_png on macOS).
    Returns a list of per-sheet error strings (empty = all OK). Callers that
    require screenshots must treat a non-empty list as failure (see build_workbook).
    """
    errors = []
    config.PNG_DIR.mkdir(parents=True, exist_ok=True)
    with xw.App(visible=True, add_book=False) as app:
        app.display_alerts = False
        wb = app.books.open(str(workbook_path))
        try:
            time.sleep(ExcelPngConst.WORKBOOK_SETTLE_SEC)
            for spec in PIVOT_SPECS:
                sheet = wb.sheets[spec["sheet"]]
                png = config.PNG_DIR / f"{sheet.name}.png"
                try:
                    sheet.activate()
                    time.sleep(ExcelPngConst.SHEET_SETTLE_SEC)
                    _capture_range_to_png(sheet.pivot_tables[0].range, png)
                except Exception as exc:  # noqa: BLE001 — Excel COM varies; collect per sheet
                    errors.append(f"{sheet.name}: {type(exc).__name__}: {exc}")
        finally:
            # Cleanup must not mask capture errors already collected above.
            try:
                wb.close()
            except Exception as exc:  # noqa: BLE001
                logger.debug("Workbook close during PNG export cleanup failed: %s", exc)
    return errors
