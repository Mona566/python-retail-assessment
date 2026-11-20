"""Summary sheet Metric/Value frame for Retail_Analysis.xlsx."""

from __future__ import annotations

import pandas as pd

from src.analysis import AnalysisBundle
from src.constants import Artifacts
from src.metrics import describe_contract


def _summary_kpi_rows(bundle: AnalysisBundle, audit: dict) -> list[tuple[str, object]]:
    """Core KPI / note rows for the Summary sheet (before Top-N list expansions)."""
    s = bundle.summary
    return [
        ("Metrics contract", describe_contract().splitlines()[0]),
        ("Gross Revenue (GBP)", s["gross_revenue"]),
        ("Return lines (is_return_line)", s.get("return_line_rows")),
        ("Confirmed return rows (Price>0 subtype)", s.get("confirmed_return_rows")),
        ("Confirmed return value (GBP)", s["return_value"]),
        ("Confirmed returned units", s.get("returned_units")),
        ("Stock adjustment rows (subtype)", s.get("stock_adjustment_rows")),
        ("Stock adjustment units", s.get("adjustment_units")),
        ("Cancellation Value abs (GBP)", s.get("cancellation_value")),
        ("Cancellation rows", s.get("cancellation_rows")),
        (
            "Returns note",
            bundle.narrative.get(
                "returns_note",
                "is_return_line = non-C Qty<0; split into is_confirmed_return (Price>0) "
                "vs is_stock_adjustment. Only confirmed return_value reduces Net.",
            ),
        ),
        ("Changed-records audit", "audit/changed_records.csv"),
        ("Net Revenue (GBP)", s["net_revenue"]),
        (
            "Net Revenue after cancellations (GBP)",
            s.get("net_revenue_after_cancellations"),
        ),
        (
            "Net note",
            (
                "Primary Net = Gross − confirmed return_value. "
                "Net after cancellations also subtracts cancellation_value "
                f"(≈ £{float(s.get('cancellation_value') or 0) / 1e6:.2f}m), "
                f"yielding ≈ £{float(s.get('net_revenue_after_cancellations') or 0) / 1e6:.1f}m."
            ),
        ),
        ("Revenue reconciliation file", "audit/revenue_reconciliation.json"),
        ("Data quality file", "audit/data_quality_summary.json"),
        ("Distinct Invoices", s["distinct_invoices"]),
        ("Distinct Customers", s["distinct_customers"]),
        ("Top Product", f"{s['top_product_stockcode']} — {s['top_product_description']}"),
        ("Top Product Revenue", s["top_product_revenue"]),
        ("Top Non-UK Country", s["top_non_uk_country"]),
        ("Top Non-UK Country Revenue", s["top_non_uk_country_revenue"]),
        ("Best Complete Month", s["best_complete_month"]),
        ("Worst Complete Month", s["worst_complete_month"]),
        ("Peak Weekday (total revenue)", s["peak_weekday"]),
        ("Peak Hour (total revenue)", s["peak_hour"]),
        ("Busiest day-hour cell", s.get("peak_day_hour_cell")),
        ("Returns min units threshold", s.get("returns_min_gross_units")),
        ("Clean_Data rows", s["rows_in_clean_data"]),
        ("Raw rows (pre-dedupe)", audit.get("raw_rows")),
        ("Removed cross-sheet clones", audit.get("removed_cross_sheet_clones")),
        ("Missing Customer ID revenue share", s["missing_customer_id_revenue_share"]),
        ("Incomplete months", ", ".join(s.get("incomplete_months") or []) or "none"),
        ("Gross by source sheet", s.get("gross_revenue_by_source_sheet")),
        ("Monthly trend note", bundle.narrative["monthly_trend"]),
        ("Missing Customer ID note", bundle.narrative["missing_customer_note"]),
        (
            "Country missing-ID note",
            bundle.narrative.get("country_missing_id_note", ""),
        ),
        ("Anomaly method note", bundle.narrative.get("anomaly_note", "")),
        ("Anomaly thresholds file", f"audit/{Artifacts.ANOMALY_THRESHOLDS_JSON}"),
        ("Anomaly thresholds (summary)", s.get("anomaly_thresholds")),
        (
            "Pivot note",
            "All six Pivot_* sheets are native Excel PivotTables (xlwings). "
            "Native PivotTables use compact Python rollups on hidden source sheets. "
            "Clean_Data retains the transaction ledger. Rebuild in Python to update the rollups. "
            "Workbook is large (~Clean_Data 1M rows); email attachments use EmailSafe ZIP "
            "(xlsx omitted) — full ZIP keeps the workbook.",
        ),
        (
            "Build note",
            "Requires local Microsoft Excel + xlwings (macOS/Windows). "
            "Phase A writes tables via xlsxwriter; Phase B creates native PivotTables.",
        ),
    ]


def _summary_topn_rows(bundle: AnalysisBundle) -> list[tuple[str, str]]:
    """Expand narrative Top-N frames into Metric/Value rows."""
    rows: list[tuple[str, str]] = []
    for _, r in bundle.narrative["top_products"].iterrows():
        rows.append(
            (
                f"Product Top {int(r['rank'])}",
                f"{r['StockCode']} | {r['Description']} | £{r['gross_revenue']:,.2f} | "
                f"units={int(r['units_sold'])} | invoices={int(r['distinct_invoices'])}",
            )
        )
    for _, r in bundle.narrative["top_countries"].iterrows():
        miss_share = r.get("missing_customer_id_revenue_share")
        miss_txt = (
            f" | missing-ID rev={float(miss_share):.0%}"
            if miss_share is not None and not pd.isna(miss_share)
            else ""
        )
        rows.append(
            (
                f"Country Top {int(r['rank'])}",
                f"{r['Country']} | £{r['gross_revenue']:,.2f} | "
                f"invoices={int(r['distinct_invoices'])} | known customers={int(r['customers'])} | "
                f"AOV=£{r['aov']:,.2f}{miss_txt}",
            )
        )
    for _, r in bundle.narrative["top_customers"].iterrows():
        rows.append(
            (
                f"Customer Top {int(r['rank'])}",
                f"ID={int(r['Customer ID']) if float(r['Customer ID']).is_integer() else r['Customer ID']} | "
                f"£{r['gross_revenue']:,.2f} | invoices={int(r['distinct_invoices'])} | "
                f"AOV=£{r['aov']:,.2f}",
            )
        )
    for _, r in bundle.narrative["top_returns"].iterrows():
        rows.append(
            (
                f"Returns Top {int(r['rank'])}",
                f"{r['StockCode']} | negative-quantity line rate={float(r['return_line_rate']):.2%} | "
                f"units_sold={int(r['units_sold'])} | return_line_units={float(r['return_line_units']):.0f} | "
                f"confirmed={float(r['returned_units']):.0f} | adj={float(r['adjustment_units']):.0f}",
            )
        )
    return rows


def _summary_frame(bundle: AnalysisBundle, audit: dict) -> pd.DataFrame:
    """Build the Summary sheet as Metric/Value rows (KPIs, notes, Top-N lists)."""
    rows = _summary_kpi_rows(bundle, audit) + _summary_topn_rows(bundle)
    return pd.DataFrame(rows, columns=["Metric", "Value"])
