"""Sanity checks for the versioned metrics contract."""

from __future__ import annotations

from src.config import ATTACHMENT_LIMIT_BYTES, ATTACHMENT_LIMIT_MB
from src.metrics import (
    CONTRACT,
    METRICS_VERSION,
    REVENUE_MATCH_TOLERANCE_GBP,
    RETURNS_MIN_GROSS_UNITS,
    describe_contract,
)


def test_metrics_version_and_thresholds():
    assert METRICS_VERSION == CONTRACT.version == "1.6.2"
    assert RETURNS_MIN_GROSS_UNITS == 50
    assert REVENUE_MATCH_TOLERANCE_GBP == 0.05
    assert ATTACHMENT_LIMIT_BYTES == 18_000_000
    assert ATTACHMENT_LIMIT_MB == 18.0
    assert CONTRACT.product_top_n == 10
    assert CONTRACT.country_top_n == 5
    assert CONTRACT.customer_top_n == 20
    assert CONTRACT.returns_top_n == 10
    assert CONTRACT.missing_customer_country_top_n == 5
    text = describe_contract()
    assert "Gross Revenue" in text
    assert "Net Revenue" in text
    assert "Return Rate" in text
    assert "Confirmed return" in text
    assert "Stock adjustments" in text
    assert f"top {CONTRACT.product_top_n}" in text
    assert f"Returns top {CONTRACT.returns_top_n}" in text
