"""Executive summary metrics and narrative prose."""

from __future__ import annotations

import pandas as pd

from src.analysis_types import _active_sales
from src.constants import Artifacts
from src.metrics import CONTRACT, RETURNS_MIN_GROSS_UNITS


def _revenue_customer_metrics(df: pd.DataFrame, sales: pd.DataFrame) -> dict:
    """Gross / return / cancellation / customer-coverage metrics from the ledger."""
    cancellations = df.loc[df["is_cancellation"]]
    gross = float(sales["gross_revenue"].sum())
    returns = float(sales["return_value"].sum())
    cancellation_value = float((-cancellations["line_revenue"]).clip(lower=0).sum())
    returned_units = float(sales["returned_units"].sum())
    adjustment_units = (
        float(df["adjustment_units"].sum()) if "adjustment_units" in df.columns else 0.0
    )
    stock_adj_rows = (
        int(df["is_stock_adjustment"].sum()) if "is_stock_adjustment" in df.columns else 0
    )
    confirmed_return_rows = (
        int(df["is_confirmed_return"].sum()) if "is_confirmed_return" in df.columns else 0
    )
    return_line_rows = (
        int(df["is_return_line"].sum()) if "is_return_line" in df.columns else 0
    )
    net = gross - returns
    net_after_cancellations = gross - returns - cancellation_value
    missing_mask = sales["Customer ID"].isna()
    missing_rev = float(sales.loc[missing_mask, "gross_revenue"].sum())
    missing_share = missing_rev / gross if gross else 0.0
    missing_by_country = (
        sales.loc[missing_mask]
        .groupby(sales.loc[missing_mask, "Country"].fillna("Unknown"))["gross_revenue"]
        .sum()
        .sort_values(ascending=False)
        .head(CONTRACT.missing_customer_country_top_n)
    )
    sheet_rev = (
        sales.groupby(sales["source_sheet"].astype(str))["gross_revenue"].sum()
        if "source_sheet" in sales.columns
        else pd.Series(dtype=float)
    )
    return {
        "gross_revenue": gross,
        "return_value": returns,
        "returned_units": returned_units,
        "adjustment_units": adjustment_units,
        "return_line_rows": return_line_rows,
        "stock_adjustment_rows": stock_adj_rows,
        "confirmed_return_rows": confirmed_return_rows,
        "cancellation_value": cancellation_value,
        "cancellation_rows": int(len(cancellations)),
        "net_revenue": net,
        "net_revenue_after_cancellations": net_after_cancellations,
        "email_total_revenue": net,
        "distinct_invoices": int(sales["Invoice"].nunique()),
        "distinct_customers": int(sales["Customer ID"].nunique()),
        "missing_customer_id_revenue": missing_rev,
        "missing_customer_id_revenue_share": missing_share,
        "missing_customer_id_revenue_by_country_top5": {
            str(k): float(v) for k, v in missing_by_country.items()
        },
        "gross_revenue_by_source_sheet": {str(k): float(v) for k, v in sheet_rev.items()},
    }


def _tops_peaks_metrics(sales: pd.DataFrame, analyses: dict) -> dict:
    """Top product/country, complete-month extremes, and peak day/hour/cell."""
    product = analyses["product"]
    country = analyses["country"]
    monthly = analyses["monthly"]
    timing = analyses["timing"]

    top_product = product.iloc[0] if len(product) else None
    top_country = country.iloc[0] if len(country) else None
    complete = monthly.loc[monthly["is_complete_month"]]
    if len(complete):
        best_month_row = complete.loc[complete["gross_revenue"].idxmax()]
        worst_month_row = complete.loc[complete["gross_revenue"].idxmin()]
        best_month = str(best_month_row["year_month"])
        worst_month = str(worst_month_row["year_month"])
    else:
        best_month = worst_month = None

    if len(sales):
        peak_day = str(sales.groupby("weekday_label")["gross_revenue"].sum().idxmax())
        peak_hour = int(sales.groupby("hour")["gross_revenue"].sum().idxmax())
    else:
        peak_day = peak_hour = None
    peak_cell = None
    if len(timing):
        cell = timing.loc[timing["gross_revenue"].idxmax()]
        peak_cell = f"{cell['weekday_label']} @ {int(cell['hour'])}:00"

    incomplete = (
        monthly.loc[~monthly["is_complete_month"], "year_month"].astype(str).tolist()
        if len(monthly)
        else []
    )
    return {
        "top_product_stockcode": None if top_product is None else str(top_product["StockCode"]),
        "top_product_description": None
        if top_product is None
        else str(top_product["Description"]),
        "top_product_revenue": None
        if top_product is None
        else float(top_product["gross_revenue"]),
        "top_non_uk_country": None if top_country is None else str(top_country["Country"]),
        "top_non_uk_country_revenue": None
        if top_country is None
        else float(top_country["gross_revenue"]),
        "best_complete_month": best_month,
        "worst_complete_month": worst_month,
        "peak_weekday": peak_day,
        "peak_hour": peak_hour,
        "peak_day_hour_cell": peak_cell,
        "incomplete_months": incomplete,
    }


def _compute_summary_metrics(df: pd.DataFrame, analyses: dict) -> dict:
    """
    Compute executive Summary metrics (gross/net, peaks, tops, incomplete months).

    Expects ``analyses`` keys: product, country, monthly, timing (and optional others).
    """
    sales = _active_sales(df)
    return {
        "metrics_contract_version": CONTRACT.version,
        **_revenue_customer_metrics(df, sales),
        **_tops_peaks_metrics(sales, analyses),
        "returns_min_gross_units": RETURNS_MIN_GROSS_UNITS,
        "rows_in_clean_data": int(len(df)),
    }


def _country_missing_id_bits(country: pd.DataFrame) -> list[str]:
    """Per-country missing-ID share snippets for the narrative note."""
    bits: list[str] = []
    for _, row in country.head(CONTRACT.country_top_n).iterrows():
        share = row.get("missing_customer_id_revenue_share")
        if share is not None and not pd.isna(share) and float(share) > 0:
            bits.append(
                f"{row['Country']}: missing-ID {float(share):.0%} of country revenue "
                f"(£{float(row['missing_customer_id_revenue']):,.0f}); "
                f"known customers={int(row['customers'])}"
            )
    return bits


def _anomaly_narrative_note(anomaly_meta: dict) -> str:
    """Prose describing anomaly fences / quotas for the Summary sheet."""
    if not anomaly_meta:
        return (
            f"Anomalies: IQR×3 fences + rules; see audit/{Artifacts.ANOMALY_THRESHOLDS_JSON}."
        )
    q = anomaly_meta.get("quota") or {}
    return (
        f"Anomalies: IQR k={anomaly_meta.get('iqr_k')} on positive non-C sales → "
        f"qty_hi={anomaly_meta.get('qty_hi')}, price_hi={anomaly_meta.get('price_hi')}, "
        f"line_hi={anomaly_meta.get('line_hi')}; customer invoice_count_hi="
        f"{anomaly_meta.get('customer_invoice_count_hi')} "
        f"({anomaly_meta.get('customers_above_frequency_fence')} customers above fence). "
        f"Top-{anomaly_meta.get('top_n')} quota monetary/adjustment/customer = "
        f"{q.get('monetary')}/{q.get('stock_adjustment')}/{q.get('customer_frequency')}; "
        f"selected {anomaly_meta.get('selected_bucket_counts')}. "
        f"Details: {anomaly_meta.get('audit_file', f'audit/{Artifacts.ANOMALY_THRESHOLDS_JSON}')}."
    )


def _build_narrative(
    df: pd.DataFrame, analyses: dict, summary: dict, anomaly_meta: dict
) -> dict:
    """
    Build prose notes and Top-N DataFrames for Summary sheet / email context.

    Returns keys such as monthly_trend, missing_customer_note, returns_note,
    anomaly_note, and top_products / top_countries / top_customers / top_returns.
    """
    product = analyses["product"]
    country = analyses["country"]
    monthly = analyses["monthly"]

    complete_n = int(monthly["is_complete_month"].sum()) if len(monthly) else 0
    incomplete = summary.get("incomplete_months") or []
    sheet_rev = summary.get("gross_revenue_by_source_sheet") or {}
    sheet_bits = ", ".join(f"{k}=£{v:,.0f}" for k, v in sheet_rev.items())
    missing_by_country = summary.get("missing_customer_id_revenue_by_country_top5") or {}
    missing_country_note = ", ".join(
        f"{idx}=£{val:,.0f}" for idx, val in missing_by_country.items()
    )
    country_miss_bits = _country_missing_id_bits(country)

    return {
        "monthly_trend": (
            f"Across both worksheets ({complete_n} complete months by date-span rule), "
            f"revenue peaks in {summary.get('best_complete_month')} and is lowest in "
            f"{summary.get('worst_complete_month')}. "
            f"Dataset span {df['InvoiceDate'].min().date() if df['InvoiceDate'].notna().any() else 'n/a'} → "
            f"{df['InvoiceDate'].max().date() if df['InvoiceDate'].notna().any() else 'n/a'}; "
            f"incomplete edge months excluded from high/low: {', '.join(incomplete) or 'none'}. "
            f"Gross by source sheet: {sheet_bits or 'n/a'}. "
            "Novembers are seasonally strong in both years."
        ),
        "missing_customer_note": (
            f"Missing Customer ID accounts for "
            f"{summary['missing_customer_id_revenue_share']:.1%} of gross revenue "
            f"(£{summary['missing_customer_id_revenue']:,.2f}); remain in totals, "
            f"excluded from Customer Top 20. "
            f"Top countries by missing-ID revenue: {missing_country_note or 'n/a'}."
        ),
        "country_missing_id_note": (
            "Country customer counts are known IDs only; revenue still includes anonymous lines. "
            + (
                " | ".join(country_miss_bits)
                if country_miss_bits
                else "No material missing-ID share in Top countries."
            )
        ),
        "returns_note": (
            f"Return indicator (is_return_line = non-C & Qty<0): "
            f"{summary['return_line_rows']:,} rows = "
            f"confirmed_return {summary['confirmed_return_rows']:,} + "
            f"stock_adjustment {summary['stock_adjustment_rows']:,}. "
            f"Confirmed (Price>0): units={summary['returned_units']:,.0f}, "
            f"value=£{summary['return_value']:,.2f} (deducted from Net). "
            f"Stock adjustments (typically Price<=0): "
            f"units={summary['adjustment_units']:,.0f} (NOT deducted from Net). "
            f"Cancellation abs value=£{summary['cancellation_value']:,.2f} "
            f"across {summary['cancellation_rows']:,} rows. "
            f"Returns Top-N: units_sold>={RETURNS_MIN_GROSS_UNITS}; "
            f"rank by return_line_rate=(returned+adjustment)/units_sold "
            f"(subtypes still shown as return_rate / adjustment_rate)."
        ),
        "top_products": product.head(CONTRACT.product_top_n)[
            [
                "rank",
                "StockCode",
                "Description",
                "gross_revenue",
                "units_sold",
                "distinct_invoices",
            ]
        ],
        "top_countries": country.head(CONTRACT.country_top_n),
        "top_customers": analyses["customer"].head(CONTRACT.customer_top_n),
        "top_returns": analyses["returns"].head(CONTRACT.returns_top_n),
        "anomaly_note": _anomaly_narrative_note(anomaly_meta),
    }


def build_summary(df: pd.DataFrame, analyses: dict) -> tuple[dict, dict]:
    """Build executive ``summary`` metrics dict and prose ``narrative`` for Summary/email."""
    summary = _compute_summary_metrics(df, analyses)
    anomaly_meta = analyses.get("anomaly_meta") or {}
    narrative = _build_narrative(df, analyses, summary, anomaly_meta)
    return summary, narrative
