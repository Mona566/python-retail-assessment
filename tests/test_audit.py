"""Tests for data-quality and revenue reconciliation audit builders."""

from __future__ import annotations

from src.data_processing import (
    apply_cross_sheet_dedupe,
    build_data_quality_summary,
    build_revenue_reconciliation,
    enrich,
)


def test_quality_and_revenue_audit(raw_transactions):
    kept, removed, dedupe = apply_cross_sheet_dedupe(raw_transactions)
    clean = enrich(kept)
    audit = {
        "raw_rows": len(raw_transactions),
        "clean_rows": len(kept),
        "removed_cross_sheet_clones": len(removed),
        "dedupe": dedupe,
    }
    quality = build_data_quality_summary(clean, audit)
    revenue = build_revenue_reconciliation(clean)

    assert quality["row_reconciliation"]["identity_check_raw_eq_clean_plus_removed"]
    assert quality["dedupe"]["matches_multiset_max"] is True
    assert "return_lines" in quality["classification_counts"]
    assert "confirmed_returns" in quality["classification_counts"]
    assert "stock_adjustments" in quality["classification_counts"]
    assert quality["classification_counts"]["stock_adjustments"] >= 1
    assert quality["return_line_taxonomy"][
        "identity_return_lines_eq_confirmed_plus_adjustments"
    ]
    assert quality["classification_counts"]["return_lines"] == (
        quality["classification_counts"]["confirmed_returns"]
        + quality["classification_counts"]["stock_adjustments"]
    )
    assert revenue["identities"]["net_eq_gross_minus_confirmed_returns"]
    assert revenue["identities"]["net_after_c_eq_net_minus_cancellation"]
