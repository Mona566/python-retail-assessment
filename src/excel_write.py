"""Phase A: prepare Clean_Data and write base workbook via xlsxwriter."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.analysis import AnalysisBundle
from src.constants import Artifacts, Sheets
from src.data_processing import CLEAN_DATA_COLUMNS
from src.excel_specs import (
    CLEAN_DATA_EXCEL_RENAME,
    PIVOT_SPECS,
    RETURNS_SOURCE_EXCEL_RENAME,
)
from src.excel_summary import _summary_frame
from src.logging_utils import get_logger
from src.metrics import CONTRACT

logger = get_logger()


def _as_excel_text(value) -> str:
    """
    Stringify IDs / YYYY-MM labels for Excel.

    Do NOT prefix a visible apostrophe — xlsxwriter already writes Python str as
    text cells. A leading ' would appear literally in PivotTables/PNGs.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value)
    if text.lower() in {"nan", "<na>", "none"}:
        return ""
    # Strip a previously-baked leading apostrophe if present in source.
    if text.startswith("'") and len(text) > 1:
        text = text[1:]
    return text


def _prepare_clean_for_excel(clean: pd.DataFrame) -> pd.DataFrame:
    """Select CLEAN_DATA_COLUMNS and coerce IDs/flags to Excel-friendly text/int types."""
    missing = [c for c in CLEAN_DATA_COLUMNS if c not in clean.columns]
    if missing:
        raise KeyError(f"Clean data missing columns required for Excel: {missing}")
    out = clean[CLEAN_DATA_COLUMNS].copy()

    # Keep codes/labels as plain strings (xlsxwriter @-style text, no visible ').
    for col in (
        "year_month",
        "weekday_label",
        "StockCode",
        "Invoice",
        "product_label",
        "Description",
        "Description_std",
    ):
        if col in out.columns:
            out[col] = out[col].map(_as_excel_text)
    out["Country"] = out["Country"].astype("string").fillna("")
    if "Country_NonUK" in out.columns:
        out["Country_NonUK"] = out["Country_NonUK"].astype("string").fillna("")
    if "Description_std" in out.columns:
        out["Description_std"] = out["Description_std"].astype("string").fillna("")

    if "Customer ID" in out.columns:
        cid = pd.to_numeric(out["Customer ID"], errors="coerce")
        # Integer IDs as plain digits; non-integer / NA → "" or original string.
        as_int = cid.notna() & (cid == cid.round(0))
        formatted = pd.Series("", index=out.index, dtype="object")
        formatted.loc[as_int] = cid.loc[as_int].astype("int64").astype(str)
        other = out["Customer ID"].notna() & ~as_int
        formatted.loc[other] = out.loc[other, "Customer ID"].astype(str)
        out["Customer ID"] = formatted

    bool_cols = [
        c
        for c in out.columns
        if c.startswith("is_") or c.endswith("_key") or c == "exact_dup_within_sheet"
    ]
    for c in bool_cols:
        out[c] = out[c].fillna(False).astype("int8")

    flag_cols = [c for c in out.columns if "flag" in c]
    for c in flag_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0).astype("int8")

    return out


def _publication_sources(bundle: AnalysisBundle) -> dict[str, pd.DataFrame]:
    """Compact Top-N / rollup frames written as hidden Pivot source sheets."""
    def ranked_labels(frame: pd.DataFrame, labels: pd.Series) -> pd.Series:
        return frame["rank"].map(lambda rank: f"{rank:02d}") + " | " + labels

    product = bundle.product.head(CONTRACT.product_top_n)
    country = bundle.country.head(CONTRACT.country_top_n)
    customer = bundle.customer.head(CONTRACT.customer_top_n)
    returns = bundle.returns.head(CONTRACT.returns_top_n).copy()
    returns["product_label"] = ranked_labels(
        returns, returns["StockCode"] + " — " + returns["Description"].fillna("")
    )
    returns = returns[list(RETURNS_SOURCE_EXCEL_RENAME)].rename(columns=RETURNS_SOURCE_EXCEL_RENAME)
    monthly = bundle.monthly.rename(columns={
        "year_month": "Year-month", "gross_revenue": "Revenue (GBP)",
        "distinct_invoices": "Monthly distinct invoices", "customers": "Monthly known customers",
        "is_complete_month": "Is complete month",
    })
    return {
        "Product_Source": pd.DataFrame({
            "Product": ranked_labels(product, product["StockCode"] + " — " + product["Description"].fillna("")),
            "Product revenue (GBP)": product["gross_revenue"],
            "Product units sold": product["units_sold"],
            "Distinct invoices": product["distinct_invoices"],
        }),
        "Country_Source": pd.DataFrame({
            "Country (excl. UK)": ranked_labels(country, country["Country"]),
            "Non-UK revenue (GBP)": country["gross_revenue"],
            "Non-UK distinct invoices": country["distinct_invoices"],
            "Non-UK known customers": country["customers"],
            "Average order value (GBP)": country["aov"],
        }),
        "Monthly_Source": monthly,
        "Returns_Source": returns,
        "Customer_Source": pd.DataFrame({
            "Customer ID": customer["Customer ID"].map(_as_excel_text),
            "Revenue rank": customer["rank"],
            "Known-customer revenue (GBP)": customer["gross_revenue"],
            "Customer distinct invoices": customer["distinct_invoices"],
            "Average order value (GBP)": customer["aov"],
        }),
        "Timing_Source": bundle.timing.rename(columns={
            "weekday_label": "Day of week", "hour": "Hour", "gross_revenue": "Revenue (GBP)",
        }),
    }


def _apply_clean_column_formats(writer: pd.ExcelWriter, clean_out: pd.DataFrame) -> None:
    """Explicit text/money/qty formats for Clean_Data (avoids date/number coercion)."""
    text_fmt = writer.book.add_format({"num_format": "@"})
    money_fmt = writer.book.add_format({"num_format": "£#,##0.00"})
    qty_fmt = writer.book.add_format({"num_format": "#,##0"})
    ws_clean = writer.sheets[Sheets.CLEAN_DATA.value]
    headers = list(clean_out.columns)
    for col_name in (
        "Invoice",
        "StockCode",
        "StockCode_raw",
        "Year-month",
        "Day of week",
        "Customer ID",
        "Product",
        "Description",
        "Description_std",
        "Country (excl. UK)",
        "return_line_class",
    ):
        if col_name in headers:
            col_idx = headers.index(col_name)
            ws_clean.set_column(col_idx, col_idx, None, text_fmt)
    for col_name in (
        "Line revenue (GBP)",
        "Revenue (GBP)",
        "Confirmed return value (GBP)",
        "Cancellation value (GBP)",
        "Product revenue (GBP)",
        "Non-UK revenue (GBP)",
        "Known-customer revenue (GBP)",
        "Price",
    ):
        if col_name in headers:
            col_idx = headers.index(col_name)
            ws_clean.set_column(col_idx, col_idx, 14, money_fmt)
    for col_name in (
        "Quantity",
        "Gross units",
        "Confirmed returned units",
        "Stock adjustment units",
        "Cancelled units",
        "Product units sold",
    ):
        if col_name in headers:
            col_idx = headers.index(col_name)
            ws_clean.set_column(col_idx, col_idx, 12, qty_fmt)


def _write_base_workbook_xlsxwriter(
    clean_out: pd.DataFrame,
    bundle: AnalysisBundle,
    audit: dict,
    output_path: Path,
) -> None:
    """
    Phase A: write Clean_Data, Summary, Anomalies, hidden sources, and empty Pivot sheets
    via pandas + XlsxWriter (no Excel COM required for this step).
    """
    summary_df = _summary_frame(bundle, audit)
    anom = bundle.anomalies.copy()
    if "InvoiceDate" in anom.columns:
        anom["InvoiceDate"] = anom["InvoiceDate"].astype(str)
    if anom.empty:
        anom = pd.DataFrame({"note": ["No anomalies flagged"]})

    sources = _publication_sources(bundle)
    clean_out = clean_out.rename(columns=CLEAN_DATA_EXCEL_RENAME)

    logger.info(
        "Writing base workbook | engine=xlsxwriter rows=%s cols=%s",
        f"{len(clean_out):,}",
        clean_out.shape[1],
    )
    # xlsxwriter is far faster than openpyxl for ~1M-row Clean_Data dumps.
    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False, "strings_to_formulas": False}},
    ) as writer:
        clean_out.to_excel(writer, sheet_name=Sheets.CLEAN_DATA.value, index=False)
        summary_df.to_excel(writer, sheet_name=Sheets.SUMMARY.value, index=False, startrow=2)
        for name, frame in sources.items():
            frame.to_excel(writer, sheet_name=name, index=False)
            writer.sheets[name].hide()
        for spec in PIVOT_SPECS:
            pd.DataFrame().to_excel(writer, sheet_name=spec["sheet"], index=False)
        anom.to_excel(writer, sheet_name=Sheets.ANOMALIES.value, index=False, startrow=2)

        ws_sum = writer.sheets[Sheets.SUMMARY.value]
        ws_sum.write(0, 0, "Retail Analysis Summary")
        ws_anom = writer.sheets[Sheets.ANOMALIES.value]
        ws_anom.write(
            0,
            0,
            "Top 20 Anomalies — IQR×k fences + rules; quota mix "
            "(monetary / stock adjustment / customer frequency). "
            f"Thresholds: audit/{Artifacts.ANOMALY_THRESHOLDS_JSON}",
        )
        if bundle.narrative.get("anomaly_note"):
            ws_anom.write(1, 0, str(bundle.narrative["anomaly_note"])[:900])
        for spec in PIVOT_SPECS:
            writer.sheets[spec["sheet"]].write(0, 0, spec["sheet"].replace("_", " "))

        _apply_clean_column_formats(writer, clean_out)
