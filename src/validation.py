"""Validate analysis outputs and the generated workbook."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import xlwings as xw
from PIL import Image

from src.analysis import AnalysisBundle
from src.config import PNG_DIR
from src.constants import BusinessRules, Sheets
from src.data_processing import EXCEL_MAX_DATA_ROWS
from src.excel_specs import PIVOT_SPECS
from src.logging_utils import get_logger
from src.metrics import ANOMALY_TOP_N, HOME_COUNTRY, REVENUE_MATCH_TOLERANCE_GBP

logger = get_logger()

# Derived from PIVOT_SPECS so sheet renames cannot drift from validation.
REQUIRED_PNGS = [f"{spec['sheet']}.png" for spec in PIVOT_SPECS]
REQUIRED_PIVOT_SHEETS = {spec["sheet"] for spec in PIVOT_SPECS}

# Sheet names that must exist in Retail_Analysis.xlsx (incl. Returns_Source).
REQUIRED_SHEETS = {
    Sheets.CLEAN_DATA.value,
    Sheets.SUMMARY.value,
    Sheets.ANOMALIES.value,
    Sheets.RETURNS_SOURCE.value,
    *REQUIRED_PIVOT_SHEETS,
}


def validate_prep(clean: pd.DataFrame, audit: dict) -> list[str]:
    """
    Check prep invariants: row counts, Excel row limit, return-line taxonomy, required columns.

    Returns a list of human-readable error strings (empty means OK).
    """
    errors: list[str] = []
    if audit["clean_rows"] != len(clean):
        errors.append("audit clean_rows mismatch")
    if len(clean) > EXCEL_MAX_DATA_ROWS:
        errors.append(
            f"Clean_Data rows {len(clean)} exceed Excel limit {EXCEL_MAX_DATA_ROWS}"
        )
    if audit["raw_rows"] != audit["clean_rows"] + audit["removed_cross_sheet_clones"]:
        errors.append("raw != clean + removed cross-sheet clones")
    dedupe = audit.get("dedupe") or {}
    if dedupe and dedupe.get("matches_multiset_max") is False:
        errors.append("cross-sheet dedupe does not match multiset max policy check")
    required = {
        "gross_revenue",
        "inv_flag_by_product",
        "source_sheet",
        "source_row",
        "Invoice",
        "is_return_line",
        "is_confirmed_return",
        "is_stock_adjustment",
        "return_line_class",
        "is_negative_qty",
        "StockCode_raw",
    }
    missing = required - set(clean.columns)
    if missing:
        errors.append(f"missing columns: {missing}")
    if "is_return_line" in clean.columns:
        parts = clean["is_confirmed_return"].astype(bool) | clean["is_stock_adjustment"].astype(
            bool
        )
        if not bool((clean["is_return_line"].astype(bool) == parts).all()):
            errors.append("is_return_line must equal confirmed_return XOR stock_adjustment cover")
        if bool((clean["is_confirmed_return"] & clean["is_stock_adjustment"]).any()):
            errors.append("is_confirmed_return and is_stock_adjustment overlap")
    return errors


def _validate_summary_identities(clean: pd.DataFrame, bundle: AnalysisBundle) -> list[str]:
    """Gross / invoice count / net-after-cancellations vs recomputed ledger."""
    errors: list[str] = []
    sales = clean.loc[~clean["is_cancellation"]]
    gross = float(sales["gross_revenue"].sum())
    if abs(gross - bundle.summary["gross_revenue"]) > REVENUE_MATCH_TOLERANCE_GBP:
        errors.append("summary gross_revenue mismatch")
    if bundle.summary["distinct_invoices"] != int(sales["Invoice"].nunique()):
        errors.append("summary distinct_invoices mismatch")
    canc_val = float(
        (-clean.loc[clean["is_cancellation"], "line_revenue"]).clip(lower=0).sum()
    )
    expected_net_after = (
        bundle.summary["gross_revenue"] - bundle.summary["return_value"] - canc_val
    )
    if (
        abs(expected_net_after - bundle.summary.get("net_revenue_after_cancellations", 0))
        > REVENUE_MATCH_TOLERANCE_GBP
    ):
        errors.append("net_revenue_after_cancellations mismatch")
    return errors


def _validate_anomaly_frames(bundle: AnalysisBundle) -> list[str]:
    """Anomaly Top-N size, required columns, and threshold keys on summary."""
    errors: list[str] = []
    if len(bundle.anomalies) > ANOMALY_TOP_N:
        errors.append(f"anomalies exceed top {ANOMALY_TOP_N}")
    if len(bundle.anomalies):
        for col in ("Invoice", "Customer ID", "reason", "relevant_value", "anomaly_score"):
            if col not in bundle.anomalies.columns:
                errors.append(f"anomalies missing {col}")
    if "anomaly_thresholds" in bundle.summary:
        thr = bundle.summary["anomaly_thresholds"] or {}
        for key in ("qty_hi", "price_hi", "line_hi", "iqr_k"):
            if key not in thr:
                errors.append(f"summary.anomaly_thresholds missing {key}")
    return errors


def _validate_ranking_frames(bundle: AnalysisBundle) -> list[str]:
    """Product/returns/country frame shape, sort order, and UK exclusion."""
    errors: list[str] = []
    if len(bundle.product) == 0:
        errors.append("product analysis empty")
    if len(bundle.returns) == 0:
        errors.append("returns analysis empty")
    else:
        for col in (
            "return_line_rate",
            "return_line_units",
            "return_rate",
            "adjustment_rate",
            "adjustment_units",
            "returned_units",
        ):
            if col not in bundle.returns.columns:
                errors.append(f"returns analysis missing {col}")
        if "return_line_rate" in bundle.returns.columns and len(bundle.returns) >= 2:
            rates = bundle.returns["return_line_rate"].fillna(-1).tolist()
            if rates != sorted(rates, reverse=True):
                errors.append("returns analysis not sorted by return_line_rate desc")
    if "missing_customer_id_revenue_share" not in bundle.country.columns:
        errors.append("country analysis missing missing_customer_id_revenue_share")
    if (bundle.country["Country"] == HOME_COUNTRY).any():
        errors.append(f"country analysis still contains {HOME_COUNTRY}")
    return errors


def _validate_peaks_and_helpers(clean: pd.DataFrame, bundle: AnalysisBundle) -> list[str]:
    """Peak day/hour vs recompute, and inv_flag_by_product vs Invoice nunique sample."""
    errors: list[str] = []
    sales = clean.loc[~clean["is_cancellation"]]
    peak_day = sales.groupby("weekday_label")["gross_revenue"].sum().idxmax()
    peak_hour = int(sales.groupby("hour")["gross_revenue"].sum().idxmax())
    if bundle.summary.get("peak_weekday") != str(peak_day):
        errors.append(
            f"peak_weekday mismatch: {bundle.summary.get('peak_weekday')} vs {peak_day}"
        )
    if bundle.summary.get("peak_hour") != peak_hour:
        errors.append(
            f"peak_hour mismatch: {bundle.summary.get('peak_hour')} vs {peak_hour}"
        )
    for _, row in bundle.product.head(5).iterrows():
        code = row["StockCode"]
        sub = sales.loc[sales["StockCode"] == code]
        helper = int(sub["inv_flag_by_product"].sum())
        nuniq = int(sub["Invoice"].nunique())
        if helper != nuniq:
            errors.append(f"product helper mismatch for {code}: {helper} vs {nuniq}")
    return errors


def validate_analyses(clean: pd.DataFrame, bundle: AnalysisBundle) -> list[str]:
    """
    Recompute key metrics from ``clean`` and compare against ``bundle.summary`` / frames.

    Covers gross/net identities, peak day/hour, UK exclusion, returns sort order,
    and invoice helper consistency on a sample of top products.
    """
    return (
        _validate_summary_identities(clean, bundle)
        + _validate_anomaly_frames(bundle)
        + _validate_ranking_frames(bundle)
        + _validate_peaks_and_helpers(clean, bundle)
    )


def _pivot_last_row_is_grand_total(pt) -> bool:
    """True when the PivotTable's last visible row label is 'Grand Total'."""
    values = pt.range.value
    if not values:
        return False
    last = values[-1]
    if not isinstance(last, list) or not last:
        return False
    return str(last[0]).strip().lower() == BusinessRules.PIVOT_GRAND_TOTAL_LABEL


_PIVOT_SHEETS = REQUIRED_PIVOT_SHEETS
# Pivots with distinct helpers / rates must not show a bottom Grand Total row.
_NO_BOTTOM_GRAND_TOTAL = {
    spec["sheet"] for spec in PIVOT_SPECS if spec.get("hide_bottom_grand_total")
}
_MONTHLY_LABEL_SKIP = {
    "year_month",
    "year-month",
    "values",
    BusinessRules.PIVOT_GRAND_TOTAL_LABEL,
    "none",
    "row labels",
}


def _check_pivot_presence_and_grand_totals(wb, names: set[str]) -> list[str]:
    """Require a PivotTable on each Pivot_* sheet; hide bottom GT on non-additive ones."""
    errors: list[str] = []
    for sheet in _PIVOT_SHEETS:
        if sheet in names and len(wb.sheets[sheet].pivot_tables) < 1:
            errors.append(f"no PivotTable on {sheet}")
    for sheet in _NO_BOTTOM_GRAND_TOTAL:
        if sheet not in names:
            continue
        pt = wb.sheets[sheet].pivot_tables[0]
        if _pivot_last_row_is_grand_total(pt):
            errors.append(f"non-additive Grand Total row still visible on {sheet}")
    return errors


def _check_monthly_year_month_labels(wb, names: set[str]) -> list[str]:
    """year_month row labels must be YYYY-MM text (not header / Excel serial)."""
    errors: list[str] = []
    if "Pivot_Monthly_Performance" not in names:
        return errors
    pt = wb.sheets["Pivot_Monthly_Performance"].pivot_tables[0]
    body = pt.range.value or []
    sample = None
    for r in body:
        if not isinstance(r, list) or not r:
            continue
        label = str(r[0]).strip()
        if not label or label.lower() in _MONTHLY_LABEL_SKIP:
            continue
        if label.startswith("Sum of"):
            continue
        sample = r[0]
        break
    if sample is None:
        errors.append("monthly pivot has no year_month data rows")
    else:
        label = str(sample).lstrip("'")
        if not (len(label) >= 7 and label[4] == "-" and label[:4].isdigit()):
            errors.append(f"unexpected year_month label on monthly pivot: {sample!r}")
    return errors


def validate_workbook(path: Path) -> list[str]:
    """
    Open the workbook in Excel (xlwings) and verify required sheets, PivotTables,
    hidden Grand Totals on non-additive pivots, and YYYY-MM monthly labels.

    Requires a local Microsoft Excel install. Returns error strings (empty = OK).
    """
    errors: list[str] = []
    if not path.exists():
        return [f"workbook missing: {path}"]
    if path.stat().st_size < 1000:
        errors.append("workbook unexpectedly small")

    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(str(path))
        names = {s.name for s in wb.sheets}
        missing = REQUIRED_SHEETS - names
        if missing:
            errors.append(f"missing sheets: {missing}")
        errors.extend(_check_pivot_presence_and_grand_totals(wb, names))
        errors.extend(_check_monthly_year_month_labels(wb, names))
    except Exception as exc:  # noqa: BLE001 — Excel COM errors become validation failures
        errors.append(f"workbook validation error: {type(exc).__name__}: {exc}")
    finally:
        # Cleanup only — do not mask validation errors collected above.
        try:
            if wb is not None:
                wb.close()
        except Exception as exc:  # noqa: BLE001
            logger.debug("Workbook close during validate cleanup failed: %s", exc)
        if app is not None:
            try:
                app.quit()
            except Exception as exc:  # noqa: BLE001
                logger.debug("Excel quit during validate cleanup failed: %s", exc)
    return errors


def validate_pngs() -> list[str]:
    """Ensure each REQUIRED_PNGS file exists, is a valid PNG, and is not a blank image."""
    errors: list[str] = []
    for name in REQUIRED_PNGS:
        path = PNG_DIR / name
        if not path.exists():
            errors.append(f"missing PNG: {name}")
            continue
        try:
            with Image.open(path) as image:
                if image.format != "PNG":
                    errors.append(f"invalid screenshot {name}: file is not a PNG")
                    continue
                image.verify()
            with Image.open(path) as image:
                rgb = image.convert("RGB")
                if all(low == high for low, high in rgb.getextrema()):
                    errors.append(f"invalid screenshot {name}: image is blank")
        except (OSError, ValueError, SyntaxError) as exc:
            # Real I/O / decode failures from Pillow — surface as validation errors.
            errors.append(f"invalid screenshot {name}: {exc}")
    return errors


def run_validations(
    clean: pd.DataFrame,
    audit: dict,
    bundle: AnalysisBundle,
    workbook: Path | None,
    check_pngs: bool = True,
) -> list[str]:
    """
    Run prep + analysis checks; optionally workbook (if path given) and PNG checks.

    Logs each failure. Returns the combined error list (empty = all passed).
    """
    errors = []
    errors.extend(validate_prep(clean, audit))
    errors.extend(validate_analyses(clean, bundle))
    if workbook is not None:
        errors.extend(validate_workbook(workbook))
    if check_pngs and workbook is not None:
        errors.extend(validate_pngs())
    if errors:
        for e in errors:
            logger.error("VALIDATION FAIL | %s", e)
    else:
        logger.info("VALIDATION OK")
    return errors
