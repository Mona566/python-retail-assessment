"""Business analyses, executive summary metrics, and anomaly detection."""

from __future__ import annotations

import json

import pandas as pd

from src.analysis_anomalies import (
    _relevant_value_for_reason,
    _select_anomaly_quota,
    detect_anomalies,
)
from src.analysis_rankings import (
    complete_months,
    country_performance,
    customer_value,
    monthly_performance,
    product_performance,
    returns_performance,
    sales_timing,
)
from src.analysis_summary import build_summary
from src.analysis_types import AnalysisBundle, _active_sales
from src.config import AUDIT_DIR, ensure_dirs
from src.constants import Artifacts
from src.data_processing import validate_enrich_schema
from src.logging_utils import get_logger, timed_phase

logger = get_logger()

__all__ = [
    "AnalysisBundle",
    "_relevant_value_for_reason",
    "_select_anomaly_quota",
    "build_summary",
    "complete_months",
    "country_performance",
    "customer_value",
    "detect_anomalies",
    "monthly_performance",
    "product_performance",
    "returns_performance",
    "run_analyses",
    "sales_timing",
]


def run_analyses(df: pd.DataFrame) -> AnalysisBundle:
    """
    Run all required analyses on an enriched clean ledger.

    Persists ``audit/anomaly_thresholds.json`` and returns a full ``AnalysisBundle``.
    """
    schema_errors = validate_enrich_schema(df)
    if schema_errors:
        raise ValueError(f"run_analyses requires enriched ledger: {schema_errors[0]}")
    logger.info("Running required analyses | rows=%s", f"{len(df):,}")
    # One shared non-cancellation view for all rankings (avoid 6× full-frame copies).
    sales = _active_sales(df)
    with timed_phase("analysis_product") as meta:
        product = product_performance(df, sales=sales)
        meta["rows"] = len(product)
    with timed_phase("analysis_country") as meta:
        country = country_performance(df, sales=sales)
        meta["rows"] = len(country)
    with timed_phase("analysis_monthly") as meta:
        monthly = monthly_performance(df, sales=sales)
        meta["rows"] = len(monthly)
    with timed_phase("analysis_returns") as meta:
        returns = returns_performance(df, sales=sales)
        meta["rows"] = len(returns)
    with timed_phase("analysis_customer") as meta:
        customer = customer_value(df, sales=sales)
        meta["rows"] = len(customer)
    with timed_phase("analysis_timing") as meta:
        timing = sales_timing(df, sales=sales)
        meta["rows"] = len(timing)
    with timed_phase("analysis_anomalies") as meta:
        anomalies, anomaly_meta = detect_anomalies(df)
        meta["flagged_top_n"] = len(anomalies)
    ensure_dirs()
    (AUDIT_DIR / Artifacts.ANOMALY_THRESHOLDS_JSON).write_text(
        json.dumps(anomaly_meta, indent=2), encoding="utf-8"
    )
    with timed_phase("analysis_summary"):
        summary, narrative = build_summary(
            df,
            {
                "product": product,
                "country": country,
                "monthly": monthly,
                "returns": returns,
                "customer": customer,
                "timing": timing,
                "anomaly_meta": anomaly_meta,
            },
        )
    summary["anomaly_thresholds"] = {
        "qty_hi": anomaly_meta.get("qty_hi"),
        "price_hi": anomaly_meta.get("price_hi"),
        "line_hi": anomaly_meta.get("line_hi"),
        "customer_invoice_count_hi": anomaly_meta.get("customer_invoice_count_hi"),
        "iqr_k": anomaly_meta.get("iqr_k"),
        "quota": anomaly_meta.get("quota"),
        "selected_bucket_counts": anomaly_meta.get("selected_bucket_counts"),
        "audit_file": anomaly_meta.get("audit_file"),
    }
    return AnalysisBundle(
        summary=summary,
        product=product,
        country=country,
        monthly=monthly,
        returns=returns,
        customer=customer,
        timing=timing,
        anomalies=anomalies,
        narrative=narrative,
    )
