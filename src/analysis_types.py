"""Shared analysis types and helpers."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class AnalysisBundle:
    """All analysis outputs for one clean ledger (Summary sheet + pivots + email)."""

    summary: dict
    product: pd.DataFrame
    country: pd.DataFrame
    monthly: pd.DataFrame
    returns: pd.DataFrame
    customer: pd.DataFrame
    timing: pd.DataFrame
    anomalies: pd.DataFrame
    narrative: dict


def _active_sales(df: pd.DataFrame) -> pd.DataFrame:
    """Non-cancellation rows (view) used as the default sales universe.

    Callers must treat the result as read-only — no in-place column writes —
    so rankings can share one filtered view without copying ~1M rows each time.
    """
    return df.loc[~df["is_cancellation"]]
