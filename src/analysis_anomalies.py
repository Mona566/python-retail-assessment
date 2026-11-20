"""Anomaly detection helpers and Top-N selection."""

from __future__ import annotations

import numpy as np
import pandas as pd

from src.analysis_types import _active_sales
from src.constants import AnomalyBucket, AnomalyReason, Artifacts
from src.metrics import (
    ANOMALY_IQR_K,
    ANOMALY_QUOTA_ADJUSTMENT,
    ANOMALY_QUOTA_CUSTOMER,
    ANOMALY_QUOTA_MONETARY,
    ANOMALY_TOP_N,
)


def _iqr_upper(series: pd.Series, k: float) -> float:
    """Upper fence Q3 + k·IQR used by anomaly detection."""
    q1 = float(series.quantile(0.25))
    q3 = float(series.quantile(0.75))
    return float(q3 + k * (q3 - q1))


def _anomaly_bucket(reason: str) -> str:
    """Map a reason string to quota bucket: monetary / adjustment / customer."""
    text = str(reason)
    if AnomalyReason.CUSTOMER_FREQUENCY.value in text:
        return AnomalyBucket.CUSTOMER.value
    if AnomalyReason.STOCK_ADJUSTMENT_QTY.value in text:
        return AnomalyBucket.ADJUSTMENT.value
    return AnomalyBucket.MONETARY.value


_REASON_PRIORITY = [
    AnomalyReason.CUSTOMER_FREQUENCY.value,
    AnomalyReason.STOCK_ADJUSTMENT_QTY.value,
    AnomalyReason.LARGE_QUANTITY.value,
    AnomalyReason.HIGH_PRICE.value,
    AnomalyReason.NEGATIVE_PRICE.value,
    AnomalyReason.CONFIRMED_RETURN.value,
    AnomalyReason.CANCELLATION.value,
    AnomalyReason.LARGE_LINE_VALUE.value,
    AnomalyReason.ZERO_QUANTITY.value,
]


def _primary_reason(reason: str) -> str:
    tokens = [t for t in str(reason).split(";") if t]
    return next((p for p in _REASON_PRIORITY if p in tokens), tokens[0] if tokens else "")


def _relevant_value_for_reason(row: pd.Series) -> float:
    """Pick the value that best explains the primary anomaly reason (single-row helper)."""
    primary = _primary_reason(row.get("reason", ""))
    if primary == AnomalyReason.CUSTOMER_FREQUENCY.value:
        return float(row.get("relevant_value", 0.0) or 0.0)
    if primary == AnomalyReason.STOCK_ADJUSTMENT_QTY.value:
        return float(row.get("adjustment_units", 0.0) or 0.0)
    if primary == AnomalyReason.LARGE_QUANTITY.value:
        return float(abs(row.get("Quantity", 0.0) or 0.0))
    if primary in {
        AnomalyReason.HIGH_PRICE.value,
        AnomalyReason.NEGATIVE_PRICE.value,
    }:
        return float(row.get("Price", 0.0) or 0.0)
    if primary == AnomalyReason.CONFIRMED_RETURN.value:
        returned = float(row.get("returned_units", 0.0) or 0.0)
        return returned if returned else float(abs(row.get("line_revenue", 0.0) or 0.0))
    if primary == AnomalyReason.ZERO_QUANTITY.value:
        return 0.0
    return float(abs(row.get("line_revenue", 0.0) or 0.0))


def _relevant_values_frame(flagged: pd.DataFrame) -> pd.Series:
    """Vectorized relevant_value for a flagged anomaly frame."""
    primary = flagged["reason"].map(_primary_reason)
    abs_line = flagged["line_revenue"].abs().fillna(0.0)
    returned = flagged.get("returned_units", pd.Series(0.0, index=flagged.index)).fillna(0.0)
    adjustment = flagged.get(
        "adjustment_units", pd.Series(0.0, index=flagged.index)
    ).fillna(0.0)
    existing = flagged.get(
        "relevant_value", pd.Series(0.0, index=flagged.index)
    ).fillna(0.0)
    qty = flagged["Quantity"].abs().fillna(0.0)
    price = flagged["Price"].fillna(0.0)

    confirmed = np.where(returned.to_numpy() != 0, returned.to_numpy(), abs_line.to_numpy())
    return pd.Series(
        np.select(
            [
                primary.eq(AnomalyReason.CUSTOMER_FREQUENCY.value),
                primary.eq(AnomalyReason.STOCK_ADJUSTMENT_QTY.value),
                primary.eq(AnomalyReason.LARGE_QUANTITY.value),
                primary.isin(
                    [AnomalyReason.HIGH_PRICE.value, AnomalyReason.NEGATIVE_PRICE.value]
                ),
                primary.eq(AnomalyReason.CONFIRMED_RETURN.value),
                primary.eq(AnomalyReason.ZERO_QUANTITY.value),
            ],
            [
                existing.to_numpy(dtype=float),
                adjustment.to_numpy(dtype=float),
                qty.to_numpy(dtype=float),
                price.to_numpy(dtype=float),
                confirmed,
                0.0,
            ],
            default=abs_line.to_numpy(dtype=float),
        ),
        index=flagged.index,
    )


def _select_anomaly_quota(
    flagged: pd.DataFrame,
    *,
    top_n: int,
    quota_monetary: int,
    quota_adjustment: int,
    quota_customer: int,
) -> pd.DataFrame:
    """
    Diversified Top-N: fill bucket quotas by score, then top up from leftovers.
    """
    if top_n < 1:
        raise ValueError(f"top_n must be >= 1, got {top_n}")
    if flagged.empty:
        return flagged
    work = flagged.copy()
    work["_bucket"] = work["reason"].map(_anomaly_bucket)
    work = work.sort_values("anomaly_score", ascending=False)
    work = work.drop_duplicates(subset=["Invoice", "StockCode", "source_row"], keep="first")

    quotas = {
        AnomalyBucket.MONETARY.value: quota_monetary,
        AnomalyBucket.ADJUSTMENT.value: quota_adjustment,
        AnomalyBucket.CUSTOMER.value: quota_customer,
    }
    # Scale quotas if top_n differs from default 20.
    default_total = quota_monetary + quota_adjustment + quota_customer
    if top_n != default_total and default_total > 0:
        scale = top_n / default_total
        # Allow 0 for tiny top_n so the shrink loop cannot hang when top_n < 3.
        quotas = {k: max(0, int(round(v * scale))) for k, v in quotas.items()}
        # Fix rounding so sum == top_n (prefer monetary for leftovers).
        while sum(quotas.values()) > top_n:
            shrunk = False
            for k in (
                AnomalyBucket.CUSTOMER.value,
                AnomalyBucket.ADJUSTMENT.value,
                AnomalyBucket.MONETARY.value,
            ):
                if quotas[k] > 0 and sum(quotas.values()) > top_n:
                    quotas[k] -= 1
                    shrunk = True
            if not shrunk:
                break
        while sum(quotas.values()) < top_n:
            quotas[AnomalyBucket.MONETARY.value] += 1

    picked_idx: list = []
    for bucket, n in quotas.items():
        bucket_rows = work.loc[work["_bucket"] == bucket]
        picked_idx.extend(list(bucket_rows.head(n).index))

    picked = work.loc[picked_idx]
    if len(picked) < top_n:
        rest = work.drop(index=picked_idx, errors="ignore")
        need = top_n - len(picked)
        picked = pd.concat([picked, rest.head(need)], ignore_index=False)

    out = (
        picked.sort_values("anomaly_score", ascending=False)
        .head(top_n)
        .drop(columns=["_bucket"], errors="ignore")
    )
    return out


def _fit_iqr_fences(work: pd.DataFrame, k: float) -> tuple[float, float, float]:
    """Fit IQR×k upper fences on positive non-cancellation sales."""
    pos = work.loc[(work["Quantity"] > 0) & (~work["is_cancellation"])]
    qty_hi = max(_iqr_upper(pos["Quantity"].abs(), k), 1.0) if len(pos) else 1.0
    price_hi = max(_iqr_upper(pos["Price"].clip(lower=0), k), 1.0) if len(pos) else 1.0
    line_hi = max(_iqr_upper(pos["abs_line"], k), 1.0) if len(pos) else 1.0
    return qty_hi, price_hi, line_hi


def _flag_line_anomaly_rules(
    work: pd.DataFrame, fences: tuple[float, float, float]
) -> pd.DataFrame:
    """Apply line-level anomaly rules; set reason/score; return flagged rows."""
    qty_hi, price_hi, line_hi = fences

    large_qty = (work["Quantity"] > 0) & (work["Quantity"] > qty_hi)
    high_price = work["Price"] > price_hi
    large_line = work["abs_line"] > line_hi
    neg_price = work["Price"] < 0
    zero_qty = work["Quantity"] == 0
    if "is_confirmed_return" in work.columns:
        extreme_return = work["is_confirmed_return"] & (
            (work["abs_line"] > line_hi) | (work["returned_units"] > qty_hi)
        )
    else:
        extreme_return = pd.Series(False, index=work.index)
    if "is_stock_adjustment" in work.columns and "adjustment_units" in work.columns:
        extreme_adjust = work["is_stock_adjustment"] & (work["adjustment_units"] > qty_hi)
    else:
        extreme_adjust = pd.Series(False, index=work.index)
    extreme_cancel = work["is_cancellation"] & (work["abs_line"] > line_hi)

    parts = []
    for mask, label in [
        (large_qty, AnomalyReason.LARGE_QUANTITY.value),
        (high_price, AnomalyReason.HIGH_PRICE.value),
        (large_line, AnomalyReason.LARGE_LINE_VALUE.value),
        (neg_price, AnomalyReason.NEGATIVE_PRICE.value),
        (zero_qty, AnomalyReason.ZERO_QUANTITY.value),
        (extreme_return, AnomalyReason.CONFIRMED_RETURN.value),
        (extreme_adjust, AnomalyReason.STOCK_ADJUSTMENT_QTY.value),
        (extreme_cancel, AnomalyReason.CANCELLATION.value),
    ]:
        parts.append(pd.Series(np.where(mask, label, ""), index=work.index))
    work["reason"] = (
        parts[0]
        .str.cat(parts[1:], sep=";")
        .str.replace(r";+", ";", regex=True)
        .str.strip(";")
    )
    work["anomaly_score"] = (
        np.where(large_qty, work["Quantity"].abs() / qty_hi, 0.0)
        + np.where(high_price, work["Price"] / price_hi, 0.0)
        + np.where(large_line, work["abs_line"] / line_hi, 0.0)
        + np.where(neg_price, 5.0, 0.0)
        + np.where(zero_qty, 2.0, 0.0)
        + np.where(extreme_return, work["abs_line"] / line_hi + 1.0, 0.0)
        + np.where(
            extreme_adjust,
            1.0 + work["adjustment_units"].fillna(0) / qty_hi,
            0.0,
        )
        + np.where(extreme_cancel, work["abs_line"] / line_hi + 1.5, 0.0)
    )
    return work.loc[work["reason"] != ""].copy()


def _flag_customer_frequency(
    df: pd.DataFrame, k: float, *, quota_customer: int
) -> tuple[list[dict], float, int]:
    """Flag customers with abnormal distinct-invoice frequency."""
    sales = _active_sales(df)
    known = sales.loc[sales["Customer ID"].notna()]
    cust = known.groupby("Customer ID")["Invoice"].nunique().rename("invoice_count")
    c_hi = 1.0
    hot_n = 0
    extra_rows: list[dict] = []
    if not cust.empty:
        c_hi = max(_iqr_upper(cust.astype(float), k), 1.0)
        hot = cust[cust > c_hi].sort_values(ascending=False)
        hot_n = int(len(hot))
        # Wide candidate pool; quota selects Top customer slots.
        pool = hot.head(max(20, quota_customer * 5))
        if not pool.empty:
            samples = (
                known.loc[known["Customer ID"].isin(pool.index)]
                .drop_duplicates(subset=["Customer ID"], keep="first")
                .set_index("Customer ID")
            )
            for cid, cnt in pool.items():
                if cid not in samples.index:
                    continue
                sample = samples.loc[cid]
                extra_rows.append(
                    {
                        **sample.to_dict(),
                        "Customer ID": cid,
                        "reason": AnomalyReason.CUSTOMER_FREQUENCY.value,
                        "relevant_value": float(cnt),
                        "anomaly_score": float(cnt / c_hi) + 50.0,
                        "adjustment_units": float(sample.get("adjustment_units", 0) or 0),
                        "returned_units": float(sample.get("returned_units", 0) or 0),
                    }
                )
    return extra_rows, c_hi, hot_n


_ANOMALY_OUTPUT_COLS = [
    "Invoice",
    "Customer ID",
    "StockCode",
    "InvoiceDate",
    "Quantity",
    "Price",
    "line_revenue",
    "Country",
    "source_sheet",
    "source_row",
    "reason",
    "relevant_value",
    "anomaly_score",
]


def _build_anomaly_meta(
    *,
    k: float,
    fences: tuple[float, float, float],
    c_hi: float,
    hot_n: int,
    top_n: int,
    flagged_n: int,
    quota_monetary: int,
    quota_adjustment: int,
    quota_customer: int,
) -> dict:
    """Thresholds / method description persisted to anomaly_thresholds.json."""
    qty_hi, price_hi, line_hi = fences
    return {
        "method": (
            "IQR upper fence (Q3 + k·IQR) on positive non-cancellation sales for "
            "Quantity, Price, |line_revenue|; plus rule flags for negative price, "
            "zero qty, extreme return/adjustment/cancellation; plus customer "
            "distinct-invoice frequency IQR fence. Top-N selected with quotas."
        ),
        "iqr_k": k,
        "fit_universe": "Quantity > 0 AND NOT is_cancellation",
        "qty_hi": float(qty_hi),
        "price_hi": float(price_hi),
        "line_hi": float(line_hi),
        "customer_invoice_count_hi": float(c_hi),
        "customers_above_frequency_fence": hot_n,
        "top_n": int(top_n),
        "quota": {
            "monetary": quota_monetary,
            "stock_adjustment": quota_adjustment,
            "customer_frequency": quota_customer,
        },
        "relevant_value_policy": (
            "Primary reason → value: large_qty→|Quantity|; high/neg price→Price; "
            "stock_adjustment→adjustment_units; customer_freq→invoice_count; "
            "else → |line_revenue|."
        ),
        "flagged_rows_before_top_n": int(flagged_n),
        "audit_file": f"audit/{Artifacts.ANOMALY_THRESHOLDS_JSON}",
    }


def _finalize_anomaly_frame(
    flagged: pd.DataFrame,
    *,
    top_n: int,
    quota_monetary: int,
    quota_adjustment: int,
    quota_customer: int,
) -> pd.DataFrame:
    """Fill relevant_value, apply quota Top-N, attach rank column."""
    if "adjustment_units" not in flagged.columns:
        flagged["adjustment_units"] = 0.0
    if "returned_units" not in flagged.columns:
        flagged["returned_units"] = 0.0
    if "relevant_value" not in flagged.columns:
        flagged["relevant_value"] = np.nan
    flagged["relevant_value"] = _relevant_values_frame(flagged)

    selected = _select_anomaly_quota(
        flagged,
        top_n=top_n,
        quota_monetary=quota_monetary,
        quota_adjustment=quota_adjustment,
        quota_customer=quota_customer,
    )
    anom = selected[_ANOMALY_OUTPUT_COLS].copy().reset_index(drop=True)
    anom.insert(0, "rank", np.arange(1, len(anom) + 1))
    return anom


def _selected_bucket_counts(anom: pd.DataFrame) -> dict[str, int]:
    """Count selected rows per quota bucket for audit meta."""
    buckets = anom["reason"].map(_anomaly_bucket)
    return {
        AnomalyBucket.MONETARY.value: int(buckets.eq(AnomalyBucket.MONETARY.value).sum()),
        AnomalyBucket.ADJUSTMENT.value: int(
            buckets.eq(AnomalyBucket.ADJUSTMENT.value).sum()
        ),
        AnomalyBucket.CUSTOMER.value: int(buckets.eq(AnomalyBucket.CUSTOMER.value).sum()),
    }


def detect_anomalies(
    df: pd.DataFrame, top_n: int | None = None
) -> tuple[pd.DataFrame, dict]:
    """
    Reproducible anomaly screen (assignment Step 3).

    Method
    ------
    1. Fit IQR×k upper fences on positive non-cancellation sales for Quantity,
       Price, and |line_revenue| (k = ANOMALY_IQR_K).
    2. Flag line rules: large qty/price/line, negative price, zero qty,
       extreme confirmed return / stock adjustment / cancellation.
       (Large negative quantities are covered via cancel/adjust/return rules;
       large_qty only applies to Quantity > 0.)
    3. Flag abnormal customer purchasing: known customers whose distinct invoice
       count exceeds an IQR×k fence on invoice frequency.
    4. Score = sum of relative excesses (and small fixed boosts).
    5. relevant_value depends on the primary reason (qty / price / |line| /
       adjustment_units / invoice_count).
    6. Top-N uses quotas (monetary / adjustment / customer) so fee cancellations
       do not hide other anomaly types.

    Returns
    -------
    anomalies DataFrame (≤ top_n) and a thresholds/meta dict for audit + Summary.
    """
    if top_n is None:
        top_n = ANOMALY_TOP_N
    if not isinstance(top_n, int) or isinstance(top_n, bool) or top_n < 1:
        raise ValueError(f"top_n must be an integer >= 1, got {top_n!r}")
    if "line_revenue" not in df.columns:
        raise ValueError("detect_anomalies requires enrich column 'line_revenue'")

    # Only copy columns needed for fences/flags (avoid full-ledger duplication).
    work = df.copy()
    work["abs_line"] = work["line_revenue"].abs()
    k = float(ANOMALY_IQR_K)

    fences = _fit_iqr_fences(work, k)
    flagged = _flag_line_anomaly_rules(work, fences)

    extra_rows, c_hi, hot_n = _flag_customer_frequency(
        df, k, quota_customer=ANOMALY_QUOTA_CUSTOMER
    )
    if extra_rows:
        flagged = pd.concat([flagged, pd.DataFrame(extra_rows)], ignore_index=True)

    meta = _build_anomaly_meta(
        k=k,
        fences=fences,
        c_hi=c_hi,
        hot_n=hot_n,
        top_n=top_n,
        flagged_n=len(flagged),
        quota_monetary=ANOMALY_QUOTA_MONETARY,
        quota_adjustment=ANOMALY_QUOTA_ADJUSTMENT,
        quota_customer=ANOMALY_QUOTA_CUSTOMER,
    )

    empty = pd.DataFrame(columns=["rank", *_ANOMALY_OUTPUT_COLS])
    if flagged.empty:
        return empty, meta

    anom = _finalize_anomaly_frame(
        flagged,
        top_n=top_n,
        quota_monetary=ANOMALY_QUOTA_MONETARY,
        quota_adjustment=ANOMALY_QUOTA_ADJUSTMENT,
        quota_customer=ANOMALY_QUOTA_CUSTOMER,
    )
    meta["selected_bucket_counts"] = _selected_bucket_counts(anom)
    return anom, meta
