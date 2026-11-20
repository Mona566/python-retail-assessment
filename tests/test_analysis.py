"""Tests for src/analysis.py business metrics (synthetic clean frame)."""

from __future__ import annotations

import pandas as pd
import pytest

from src.analysis import (
    _relevant_value_for_reason,
    _select_anomaly_quota,
    build_summary,
    complete_months,
    country_performance,
    customer_value,
    detect_anomalies,
    product_performance,
    returns_performance,
    run_analyses,
    sales_timing,
)
from src.metrics import RETURNS_MIN_GROSS_UNITS


def test_complete_months_uses_date_span_not_every_calendar_day():
    df = pd.DataFrame(
        {
            "InvoiceDate": pd.to_datetime(
                ["2011-03-01", "2011-03-15", "2011-03-30", "2011-04-01", "2011-04-30"]
            ),
            "year_month": ["2011-03", "2011-03", "2011-03", "2011-04", "2011-04"],
        }
    )
    complete = complete_months(df)
    assert "2011-03" in complete and "2011-04" in complete


def test_partial_edge_month_excluded():
    df = pd.DataFrame(
        {
            "InvoiceDate": pd.to_datetime(
                ["2011-11-01", "2011-11-30", "2011-12-01", "2011-12-09"]
            ),
            "year_month": ["2011-11", "2011-11", "2011-12", "2011-12"],
        }
    )
    complete = complete_months(df)
    assert "2011-11" in complete
    assert "2011-12" not in complete


def test_country_excludes_uk_and_computes_aov(clean_transactions: pd.DataFrame):
    country = country_performance(clean_transactions)
    assert "United Kingdom" not in set(country["Country"].astype(str))
    assert "EIRE" in set(country["Country"].astype(str))
    assert "aov" in country.columns
    assert "missing_customer_id_revenue_share" in country.columns
    row = country.loc[country["Country"] == "EIRE"].iloc[0]
    assert abs(row["aov"] - row["gross_revenue"] / row["distinct_invoices"]) < 1e-9
    assert row["missing_customer_id_rows"] >= 1
    assert row["missing_customer_id_revenue"] > 0


def test_product_includes_description(clean_transactions: pd.DataFrame):
    product = product_performance(clean_transactions)
    assert "Description" in product.columns
    assert product.iloc[0]["gross_revenue"] >= product.iloc[-1]["gross_revenue"]
    # Top SKU must match recomputed gross on non-cancellation product rows
    sales = clean_transactions.loc[
        ~clean_transactions["is_cancellation"] & ~clean_transactions["is_non_product"]
    ]
    top_code = str(product.iloc[0]["StockCode"])
    expected = float(sales.loc[sales["StockCode"] == top_code, "gross_revenue"].sum())
    assert abs(float(product.iloc[0]["gross_revenue"]) - expected) < 1e-9


def test_returns_min_units_and_rate_ranking():
    rows = []
    for i in range(10):
        rows.append(_sale("LOW", units=1, returned=0, inv=f"L{i}"))
    for i in range(10):
        rows.append(_sale("LOW", units=0, returned=1, inv=f"LR{i}", price=1.0))
    for i in range(100):
        rows.append(_sale("HIGH", units=1, returned=0, inv=f"H{i}"))
    for i in range(20):
        rows.append(_sale("HIGH", units=0, returned=1, inv=f"HR{i}", price=1.0))
    for i in range(80):
        rows.append(_sale("MID", units=1, returned=0, inv=f"M{i}"))
    for i in range(40):
        rows.append(_sale("MID", units=0, returned=1, inv=f"MR{i}", price=1.0))

    out = returns_performance(pd.DataFrame(rows))
    assert (out["units_sold"] >= RETURNS_MIN_GROSS_UNITS).all()
    assert "LOW" not in set(out["StockCode"].astype(str))
    assert list(out["StockCode"].astype(str)[:2]) == ["MID", "HIGH"]
    assert out.iloc[0]["return_rate"] == 0.5
    assert out.iloc[0]["return_line_rate"] == 0.5
    assert out.iloc[0]["rank_basis"] == "return_line_rate"
    assert "cancellation_value" in out.columns
    assert "adjustment_rate" in out.columns


def test_returns_ranks_by_return_line_rate_including_adjustments():
    rows = []
    for i in range(60):
        rows.append(_sale("A", units=1, returned=0, inv=f"A{i}"))
    for i in range(30):
        rows.append(_sale("A", units=0, returned=1, inv=f"AR{i}", price=0.0))
    for i in range(60):
        rows.append(_sale("B", units=1, returned=0, inv=f"B{i}"))
    for i in range(10):
        rows.append(_sale("B", units=0, returned=1, inv=f"BR{i}", price=0.0))
    # C has confirmed returns only (lower combined rate than A)
    for i in range(60):
        rows.append(_sale("C", units=1, returned=0, inv=f"C{i}", price=1.0))
    for i in range(6):
        rows.append(_sale("C", units=0, returned=1, inv=f"CR{i}", price=2.0))
    out = returns_performance(pd.DataFrame(rows))
    assert out.iloc[0]["StockCode"] == "A"
    assert out.iloc[0]["rank_basis"] == "return_line_rate"
    assert float(out.loc[out["StockCode"] == "A", "returned_units"].iloc[0]) == 0.0
    assert float(out.loc[out["StockCode"] == "A", "return_line_rate"].iloc[0]) == 0.5
    assert list(out["return_line_rate"]) == sorted(out["return_line_rate"], reverse=True)


def test_customer_excludes_missing_id(clean_transactions: pd.DataFrame):
    cust = customer_value(clean_transactions)
    assert cust["Customer ID"].notna().all()
    assert "aov" in cust.columns


def test_peak_day_and_hour_are_separate_aggregations(clean_transactions: pd.DataFrame):
    bundle = run_analyses(clean_transactions)
    sales = clean_transactions.loc[~clean_transactions["is_cancellation"]]
    expected_day = str(sales.groupby("weekday_label")["gross_revenue"].sum().idxmax())
    expected_hour = int(sales.groupby("hour")["gross_revenue"].sum().idxmax())
    assert bundle.summary["peak_weekday"] == expected_day
    assert bundle.summary["peak_hour"] == expected_hour
    # Busiest single cell may differ from peak day / peak hour
    assert "peak_day_hour_cell" in bundle.summary


def test_net_after_cancellations(clean_transactions: pd.DataFrame, analysis_bundle):
    s = analysis_bundle.summary
    expected = s["gross_revenue"] - s["return_value"] - s["cancellation_value"]
    assert abs(s["net_revenue_after_cancellations"] - expected) < 1e-6
    assert abs(s["net_revenue"] - (s["gross_revenue"] - s["return_value"])) < 1e-6


def test_summary_net_subtracts_confirmed_returns():
    """Confirmed return (Qty<0, Price>0, non-C) must reduce Net and email_total_revenue."""
    from src.data_processing import apply_cross_sheet_dedupe, enrich

    raw = pd.DataFrame(
        [
            {
                "load_row_id": 1,
                "source_sheet": "Year 2010-2011",
                "source_row": 1,
                "Invoice": "S1",
                "StockCode": "P9",
                "Description": "ITEM",
                "Quantity": 10,
                "InvoiceDate": "2011-03-01 10:00:00",
                "Price": 5.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
            },
            {
                "load_row_id": 2,
                "source_sheet": "Year 2010-2011",
                "source_row": 2,
                "Invoice": "R1",
                "StockCode": "P9",
                "Description": "ITEM",
                "Quantity": -2,
                "InvoiceDate": "2011-03-15 10:00:00",
                "Price": 5.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
            },
            {
                "load_row_id": 3,
                "source_sheet": "Year 2010-2011",
                "source_row": 3,
                "Invoice": "S2",
                "StockCode": "P9",
                "Description": "ITEM",
                "Quantity": 10,
                "InvoiceDate": "2011-03-31 10:00:00",
                "Price": 5.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
            },
        ]
    )
    kept, _, _ = apply_cross_sheet_dedupe(raw)
    df = enrich(kept)
    assert float(df["return_value"].sum()) == 10.0
    bundle = run_analyses(df)
    s = bundle.summary
    assert s["return_value"] == 10.0
    assert s["confirmed_return_rows"] == 1
    assert abs(s["gross_revenue"] - 100.0) < 1e-9
    assert abs(s["net_revenue"] - 90.0) < 1e-9
    assert s["email_total_revenue"] == s["net_revenue"]
    assert "confirmed_return" in bundle.narrative["returns_note"]


def test_anomalies_can_flag_cancellations(clean_transactions: pd.DataFrame):
    anom, meta = detect_anomalies(clean_transactions, top_n=20)
    assert len(anom) <= 20
    assert "qty_hi" in meta and meta["iqr_k"] == 3.0
    if len(anom):
        assert "reason" in anom.columns
        assert {"Invoice", "Customer ID", "reason", "relevant_value"} <= set(anom.columns)


def test_anomalies_quota_includes_customer_and_adjustment():
    """Synthetic: monetary giant, zero-price adjust, and a hyper-active customer."""
    rows = []
    # Many one-invoice customers so the invoice-frequency IQR fence stays low
    for i in range(200):
        rows.append(
            {
                "Invoice": f"N{i}",
                "Customer ID": float(1000 + i),
                "StockCode": "N",
                "Description_std": "N",
                "InvoiceDate": "2011-04-01",
                "Quantity": 2,
                "Price": 1.0,
                "line_revenue": 2.0,
                "gross_revenue": 2.0,
                "Country": "United Kingdom",
                "source_sheet": "Year 2010-2011",
                "source_row": i + 1,
                "is_cancellation": False,
                "is_confirmed_return": False,
                "is_stock_adjustment": False,
                "is_return_line": False,
                "returned_units": 0.0,
                "adjustment_units": 0.0,
            }
        )
    # Giant sale
    rows.append(
        {
            "Invoice": "BIG",
            "Customer ID": 2.0,
            "StockCode": "X",
            "Description_std": "X",
            "InvoiceDate": "2011-04-02",
            "Quantity": 5000,
            "Price": 1.0,
            "line_revenue": 5000.0,
            "gross_revenue": 5000.0,
            "Country": "United Kingdom",
            "source_sheet": "Year 2010-2011",
            "source_row": 9001,
            "is_cancellation": False,
            "is_confirmed_return": False,
            "is_stock_adjustment": False,
            "is_return_line": False,
            "returned_units": 0.0,
            "adjustment_units": 0.0,
        }
    )
    # Extreme stock adjustment (Price=0)
    rows.append(
        {
            "Invoice": "ADJ",
            "Customer ID": None,
            "StockCode": "Y",
            "Description_std": "Y",
            "InvoiceDate": "2011-04-03",
            "Quantity": -8000,
            "Price": 0.0,
            "line_revenue": 0.0,
            "gross_revenue": 0.0,
            "Country": "United Kingdom",
            "source_sheet": "Year 2010-2011",
            "source_row": 9002,
            "is_cancellation": False,
            "is_confirmed_return": False,
            "is_stock_adjustment": True,
            "is_return_line": True,
            "returned_units": 0.0,
            "adjustment_units": 8000.0,
        }
    )
    # Hyper-active customer 99: many invoices
    for i in range(40):
        rows.append(
            {
                "Invoice": f"H{i}",
                "Customer ID": 99.0,
                "StockCode": "Z",
                "Description_std": "Z",
                "InvoiceDate": "2011-04-04",
                "Quantity": 1,
                "Price": 1.0,
                "line_revenue": 1.0,
                "gross_revenue": 1.0,
                "Country": "EIRE",
                "source_sheet": "Year 2010-2011",
                "source_row": 10000 + i,
                "is_cancellation": False,
                "is_confirmed_return": False,
                "is_stock_adjustment": False,
                "is_return_line": False,
                "returned_units": 0.0,
                "adjustment_units": 0.0,
            }
        )
    anom, meta = detect_anomalies(pd.DataFrame(rows), top_n=20)
    reasons = anom["reason"].astype(str)
    assert reasons.str.contains("unusually_large_quantity|unusually_large_line_value").any()
    assert reasons.str.contains("extreme_stock_adjustment_qty").any()
    assert reasons.str.contains("abnormal_customer_invoice_frequency").any()
    adj = anom.loc[reasons.str.contains("extreme_stock_adjustment_qty")].iloc[0]
    assert float(adj["relevant_value"]) == 8000.0
    cust = anom.loc[reasons.str.contains("abnormal_customer_invoice_frequency")].iloc[0]
    assert float(cust["relevant_value"]) == 40.0
    assert meta["selected_bucket_counts"]["customer"] >= 1
    assert meta["selected_bucket_counts"]["adjustment"] >= 1


def test_sales_timing_matrix_shape(clean_transactions: pd.DataFrame):
    timing = sales_timing(clean_transactions)
    assert {"weekday_label", "hour", "gross_revenue"} <= set(timing.columns)


def test_run_analyses_bundle_fields(analysis_bundle):
    assert analysis_bundle.summary["metrics_contract_version"].startswith("1.")
    assert len(analysis_bundle.product) >= 1
    assert len(analysis_bundle.country) >= 1


def test_complete_months_empty_and_bad_label():
    empty = pd.DataFrame({"InvoiceDate": pd.Series(dtype="datetime64[ns]"), "year_month": []})
    assert complete_months(empty) == set()
    df = pd.DataFrame(
        {
            "InvoiceDate": pd.to_datetime(["2011-01-01", "2011-01-31"]),
            "year_month": ["bad", "2011-01"],
        }
    )
    assert "2011-01" in complete_months(df)


def test_relevant_value_for_reason_branches():
    base = {
        "Quantity": -5,
        "Price": 3.5,
        "line_revenue": -10.0,
        "adjustment_units": 7.0,
        "returned_units": 0.0,
        "relevant_value": 99.0,
    }
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "abnormal_customer_invoice_frequency"})) == 99.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "extreme_stock_adjustment_qty"})) == 7.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "unusually_large_quantity"})) == 5.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "unusually_high_price"})) == 3.5
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "negative_price"})) == 3.5
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "extreme_confirmed_return", "returned_units": 4})) == 4.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "extreme_confirmed_return"})) == 10.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "zero_quantity"})) == 0.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": "extreme_cancellation"})) == 10.0
    assert _relevant_value_for_reason(pd.Series({**base, "reason": ""})) == 10.0


def test_select_anomaly_quota_scaling_and_empty():
    empty = pd.DataFrame()
    assert _select_anomaly_quota(
        empty, top_n=5, quota_monetary=12, quota_adjustment=4, quota_customer=4
    ).empty

    rows = []
    for i, (bucket_reason, score) in enumerate(
        [
            ("unusually_large_quantity", 100),
            ("unusually_high_price", 90),
            ("extreme_stock_adjustment_qty", 80),
            ("abnormal_customer_invoice_frequency", 70),
            ("unusually_large_line_value", 60),
            ("negative_price", 50),
        ]
    ):
        rows.append(
            {
                "Invoice": f"I{i}",
                "StockCode": f"S{i}",
                "source_row": i,
                "reason": bucket_reason,
                "anomaly_score": score,
            }
        )
    flagged = pd.DataFrame(rows)
    # top_n=5 scales quotas from default 20 total
    picked = _select_anomaly_quota(
        flagged, top_n=5, quota_monetary=12, quota_adjustment=4, quota_customer=4
    )
    assert len(picked) <= 5
    assert len(picked) >= 1

    # top_n=3 → rounded quotas sum > top_n (trim loop); top_n=7 → top-up loop
    assert len(
        _select_anomaly_quota(
            flagged, top_n=3, quota_monetary=12, quota_adjustment=4, quota_customer=4
        )
    ) <= 3
    assert len(
        _select_anomaly_quota(
            flagged, top_n=7, quota_monetary=12, quota_adjustment=4, quota_customer=4
        )
    ) <= 7


def test_detect_anomalies_rejects_invalid_top_n(clean_transactions: pd.DataFrame):
    with pytest.raises(ValueError, match="top_n"):
        detect_anomalies(clean_transactions, top_n=0)
    with pytest.raises(ValueError, match="top_n"):
        detect_anomalies(clean_transactions, top_n=-1)
    # Tiny positive top_n must not hang when scaling quotas
    anom, _ = detect_anomalies(clean_transactions, top_n=1)
    assert len(anom) <= 1


def test_run_analyses_requires_enrich_schema():
    with pytest.raises(ValueError, match="enriched ledger"):
        run_analyses(pd.DataFrame({"Invoice": ["1"]}))


def test_detect_anomalies_edge_empty_and_missing_cols():
    # Minimal frame with no positive sales → empty anomalies, default fences
    df = pd.DataFrame(
        [
            {
                "Invoice": "C1",
                "Customer ID": None,
                "StockCode": "X",
                "InvoiceDate": "2011-01-01",
                "Quantity": -1,
                "Price": 1.0,
                "line_revenue": -1.0,
                "Country": "EIRE",
                "source_sheet": "Y",
                "source_row": 1,
                "is_cancellation": True,
                "gross_revenue": 0.0,
                "adjustment_units": 0.0,
                "returned_units": 0.0,
            }
        ]
    )
    anom, meta = detect_anomalies(df, top_n=5)
    assert len(anom) == 0
    assert meta["qty_hi"] == 1.0

    # Missing is_confirmed_return / is_stock_adjustment → False series branches
    df2 = pd.DataFrame(
        [
            {
                "Invoice": "1",
                "Customer ID": 1.0,
                "StockCode": "X",
                "InvoiceDate": "2011-01-01",
                "Quantity": 0,
                "Price": -5.0,
                "line_revenue": 0.0,
                "Country": "EIRE",
                "source_sheet": "Y",
                "source_row": 1,
                "is_cancellation": False,
                "gross_revenue": 0.0,
                "adjustment_units": 0.0,
                "returned_units": 0.0,
            }
        ]
    )
    anom2, _ = detect_anomalies(df2, top_n=5)
    assert len(anom2) >= 1
    assert "zero_quantity" in str(anom2.iloc[0]["reason"]) or "negative_price" in str(
        anom2.iloc[0]["reason"]
    )


def test_build_summary_empty_product_monthly(clean_transactions):
    product = product_performance(clean_transactions).iloc[0:0]
    country = country_performance(clean_transactions).iloc[0:0]
    monthly = pd.DataFrame(
        columns=["year_month", "gross_revenue", "distinct_invoices", "customers", "is_complete_month"]
    )
    customer = customer_value(clean_transactions)
    timing = sales_timing(clean_transactions).iloc[0:0]
    returns = returns_performance(clean_transactions)
    # Empty sales path: use cancellations-only frame
    canc_only = clean_transactions.loc[clean_transactions["is_cancellation"]].copy()
    if len(canc_only) == 0:
        canc_only = clean_transactions.head(0)
    summary, narrative = build_summary(
        canc_only if len(canc_only) else clean_transactions.head(0),
        {
            "product": product,
            "country": country,
            "monthly": monthly,
            "returns": returns,
            "customer": customer,
            "timing": timing,
            "anomaly_meta": {},
        },
    )
    assert summary["best_complete_month"] is None
    assert "Anomalies: IQR×3" in narrative["anomaly_note"] or "IQR" in narrative["anomaly_note"]


def _sale(
    code: str, *, units: float, returned: float, inv: str, price: float | None = None
) -> dict:
    """Build a pre-enriched-like row for returns_performance aggregation tests."""
    if returned and price is None:
        price = 0.0
    if units and price is None:
        price = 1.0
    price = 0.0 if price is None else price
    is_return = returned > 0 and price > 0
    is_adj = returned > 0 and not is_return
    return {
        "StockCode": code,
        "Description_std": code,
        "gross_units": units,
        "returned_units": returned if is_return else 0.0,
        "adjustment_units": returned if is_adj else 0.0,
        "return_value": returned * price if is_return else 0.0,
        "gross_revenue": float(units) * (price if units else 0.0),
        "is_cancellation": False,
        "is_non_product": False,
        "Quantity": units if units else -returned,
        "line_revenue": (units if units else -returned) * price,
        "Invoice": inv,
    }
