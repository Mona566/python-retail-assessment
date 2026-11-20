"""Ranking and performance analyses over the clean ledger."""

from __future__ import annotations

import numpy as np
import pandas as pd

from src.analysis_types import _active_sales
from src.metrics import HOME_COUNTRY, RETURNS_MIN_GROSS_UNITS


def complete_months(df: pd.DataFrame) -> set[str]:
    """
    Months fully covered by the dataset's global date span.

    A month YYYY-MM is complete when min(InvoiceDate) <= first day of that month
    and max(InvoiceDate) >= last day of that month. Sparse trading days inside
    the month do not make it incomplete.
    """
    dates = df.loc[df["InvoiceDate"].notna(), "InvoiceDate"]
    if dates.empty:
        return set()
    dmin = dates.min()
    dmax = dates.max()
    months = (
        df.loc[df["year_month"].notna(), "year_month"].astype(str).dropna().unique()
    )
    complete: set[str] = set()
    for ym in months:
        try:
            start = pd.Timestamp(f"{ym}-01")
        except (ValueError, TypeError):
            # Skip malformed year_month labels rather than aborting the whole ranking.
            continue
        end = start + pd.offsets.MonthEnd(0)
        if dmin <= start and dmax >= end:
            complete.add(ym)
    return complete


def product_performance(
    df: pd.DataFrame, *, sales: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Rank products by gross_revenue (excludes non-product stock codes)."""
    sales = _active_sales(df) if sales is None else sales
    products = sales.loc[~sales["is_non_product"]]
    g = products.groupby("StockCode", as_index=False).agg(
        Description=("Description_std", "first"),
        gross_revenue=("gross_revenue", "sum"),
        units_sold=("gross_units", "sum"),
        distinct_invoices=("Invoice", "nunique"),
    )
    g = g.sort_values("gross_revenue", ascending=False).reset_index(drop=True)
    g["rank"] = np.arange(1, len(g) + 1)
    return g


def country_performance(
    df: pd.DataFrame, *, sales: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Non-UK country rollup with AOV and missing-Customer-ID revenue share."""
    sales = _active_sales(df) if sales is None else sales
    sales = sales.loc[sales["Country"].fillna("") != HOME_COUNTRY]
    # Precompute missing-ID revenue so one groupby covers share + row counts.
    work = sales.assign(
        _missing_id_rows=sales["Customer ID"].isna().astype("int64"),
        _missing_id_rev=sales["gross_revenue"].where(sales["Customer ID"].isna(), 0.0),
    )
    g = work.groupby("Country", as_index=False).agg(
        gross_revenue=("gross_revenue", "sum"),
        distinct_invoices=("Invoice", "nunique"),
        customers=("Customer ID", "nunique"),
        missing_customer_id_rows=("_missing_id_rows", "sum"),
        missing_customer_id_revenue=("_missing_id_rev", "sum"),
    )
    g["missing_customer_id_revenue_share"] = (
        g["missing_customer_id_revenue"] / g["gross_revenue"].replace(0, np.nan)
    )
    g["aov"] = g["gross_revenue"] / g["distinct_invoices"].replace(0, np.nan)
    g = g.sort_values("gross_revenue", ascending=False).reset_index(drop=True)
    g["rank"] = np.arange(1, len(g) + 1)
    return g


def monthly_performance(
    df: pd.DataFrame, *, sales: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Monthly revenue / invoices / customers plus ``is_complete_month`` flag."""
    sales = _active_sales(df) if sales is None else sales
    g = sales.groupby("year_month", as_index=False).agg(
        gross_revenue=("gross_revenue", "sum"),
        distinct_invoices=("Invoice", "nunique"),
        customers=("Customer ID", "nunique"),
    )
    complete = complete_months(df)
    g["is_complete_month"] = g["year_month"].isin(complete)
    g = g.sort_values("year_month").reset_index(drop=True)
    return g


def returns_performance(
    df: pd.DataFrame, *, sales: pd.DataFrame | None = None
) -> pd.DataFrame:
    """
    Return lines + cancellations by product (assignment Returns analysis).

    - return_line_units / return_line_rate: assignment-facing rate
      (confirmed returned_units + adjustment_units) / units_sold — used for Top-10 rank
    - returned_units / return_value / return_rate: is_confirmed_return subtype
    - adjustment_units / adjustment_rate: is_stock_adjustment subtype
    - cancelled_* reported separately (not in return_line_rate)
    Eligible: units_sold >= RETURNS_MIN_GROSS_UNITS.
    """
    sales = _active_sales(df) if sales is None else sales
    products = sales.loc[~sales["is_non_product"]]
    g = products.groupby("StockCode", as_index=False).agg(
        Description=("Description_std", "first"),
        units_sold=("gross_units", "sum"),
        returned_units=("returned_units", "sum"),
        adjustment_units=("adjustment_units", "sum"),
        return_value=("return_value", "sum"),
        gross_revenue=("gross_revenue", "sum"),
    )
    # Prefer enrich-time cancellation measures; fall back for lean test frames.
    canc = df.loc[df["is_cancellation"] & ~df["is_non_product"]].copy()
    if "cancelled_units" not in canc.columns:
        canc["cancelled_units"] = (-canc["Quantity"]).clip(lower=0)
    if "cancellation_value_line" not in canc.columns:
        canc["cancellation_value_line"] = (-canc["line_revenue"]).clip(lower=0)
    c = canc.groupby("StockCode", as_index=False).agg(
        cancelled_units=("cancelled_units", "sum"),
        cancellation_value=("cancellation_value_line", "sum"),
    )
    g = g.merge(c, on="StockCode", how="left")
    g["cancelled_units"] = g["cancelled_units"].fillna(0.0)
    g["cancellation_value"] = g["cancellation_value"].fillna(0.0)
    g["return_line_units"] = g["returned_units"] + g["adjustment_units"]
    g["return_rate"] = g["returned_units"] / g["units_sold"].replace(0, np.nan)
    g["adjustment_rate"] = g["adjustment_units"] / g["units_sold"].replace(0, np.nan)
    g["return_line_rate"] = g["return_line_units"] / g["units_sold"].replace(0, np.nan)
    eligible = g.loc[g["units_sold"] >= RETURNS_MIN_GROSS_UNITS].copy()
    eligible = eligible.sort_values(
        ["return_line_rate", "return_line_units", "returned_units"],
        ascending=[False, False, False],
    )
    eligible["rank_basis"] = "return_line_rate"
    eligible = eligible.reset_index(drop=True)
    eligible["rank"] = np.arange(1, len(eligible) + 1)
    return eligible


def customer_value(
    df: pd.DataFrame, *, sales: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Known-customer ranking by gross_revenue (anonymous IDs excluded)."""
    sales = _active_sales(df) if sales is None else sales
    known = sales.loc[sales["Customer ID"].notna()]
    g = known.groupby("Customer ID", as_index=False).agg(
        gross_revenue=("gross_revenue", "sum"),
        distinct_invoices=("Invoice", "nunique"),
    )
    g["aov"] = g["gross_revenue"] / g["distinct_invoices"].replace(0, np.nan)
    g = g.sort_values("gross_revenue", ascending=False).reset_index(drop=True)
    g["rank"] = np.arange(1, len(g) + 1)
    return g


def sales_timing(
    df: pd.DataFrame, *, sales: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Weekday × hour gross_revenue matrix for the Sales Timing pivot."""
    sales = _active_sales(df) if sales is None else sales
    g = (
        sales.groupby(["weekday_label", "hour"], as_index=False)["gross_revenue"]
        .sum()
        .rename(columns={"gross_revenue": "gross_revenue"})
    )
    return g.sort_values(["weekday_label", "hour"]).reset_index(drop=True)
