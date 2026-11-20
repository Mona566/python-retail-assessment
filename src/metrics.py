"""
Fixed metrics contract for the assessment.

All Summary figures, PivotTables, and the Outlook email must use these
definitions. Change only with an explicit version bump and README note.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Final

from src.constants import BusinessRules

METRICS_VERSION: Final[str] = "1.6.2"

# Anomaly detection (IQR fences + quota Top-N). Documented in detect_anomalies().
ANOMALY_IQR_K: Final[float] = 3.0
ANOMALY_TOP_N: Final[int] = 20
# Quota so monetary outliers do not crowd out adjustments / customer patterns.
ANOMALY_QUOTA_MONETARY: Final[int] = 12
ANOMALY_QUOTA_ADJUSTMENT: Final[int] = 4
ANOMALY_QUOTA_CUSTOMER: Final[int] = 4

# Re-export domain policy for a single import surface with other metrics knobs.
HOME_COUNTRY: Final[str] = BusinessRules.HOME_COUNTRY
CANCELLATION_INVOICE_PREFIX: Final[str] = BusinessRules.CANCELLATION_INVOICE_PREFIX

# Non-product stock codes excluded from Product Top 10 (still may affect totals).
# Explicit allowlist only — do not treat every ^[A-Z]+$ code as a fee (letter-only
# SKUs such as DCGSSGIRL are real catalogue items on Online Retail II).
NON_PRODUCT_STOCK_CODES: Final[frozenset[str]] = frozenset(
    {
        "POST",
        "DOT",
        "M",
        "D",
        "S",
        "B",  # Adjust bad debt
        "C2",  # Carriage
        "AMAZONFEE",
        "BANK CHARGES",
        "CRUK",
        "ADJUST",
        "ADJUST2",
    }
)

# Minimum gross units before a product enters Returns Top 10.
RETURNS_MIN_GROSS_UNITS: Final[int] = 50

# Absolute £ tolerance when reconciling summary revenue vs recomputed ledger totals.
REVENUE_MATCH_TOLERANCE_GBP: Final[float] = 0.05


@dataclass(frozen=True)
class MetricsContract:
    """
    Canonical business definitions for Summary, PivotTables, and email.

    Field values are human-readable policy strings (not computed metrics).
    Bump METRICS_VERSION + README when changing meanings.
    """

    version: str = METRICS_VERSION

    # Revenue
    gross_revenue: str = (
        "Sum of line_revenue for positive-quantity, non-cancellation sales lines "
        "(Quantity > 0 and not is_cancellation)."
    )
    return_value: str = (
        "Absolute line_revenue on confirmed returns only "
        "(is_confirmed_return: Quantity < 0, not cancellation, Price > 0). "
        "On Online Retail II this set is typically empty/near-empty."
    )
    adjustment_note: str = (
        "Within return lines, Price <= 0 rows are stock/accounting adjustments "
        "(is_stock_adjustment). They contribute adjustment_units but NOT return_value "
        "and are NOT deducted from Net."
    )
    cancellation_value: str = (
        "Sum of abs(line_revenue) on cancellation invoices (Invoice starts with C). "
        "Reported separately in Summary; not subtracted again in net_revenue."
    )
    net_revenue: str = (
        "gross_revenue - return_value (cancellations and stock adjustments excluded). "
        "Also report net_revenue_after_cancellations = gross - return_value - cancellation_value."
    )
    email_total_revenue: str = (
        "net_revenue (gross - confirmed return_value). "
        "Email also states net_revenue_after_cancellations for clarity."
    )

    # Classification (three-tier — avoid a bare ambiguous "is_return")
    is_cancellation: str = "Invoice string starts with 'C' (case-insensitive)."
    is_negative_qty: str = "Quantity < 0 (includes cancellations and return lines)."
    is_return_line: str = (
        "Assignment return indicator: Quantity < 0 AND not a cancellation. "
        "Parent bucket; always split into is_confirmed_return XOR is_stock_adjustment."
    )
    is_confirmed_return: str = (
        "Subtype of is_return_line with Price > 0 (paid customer-style return). "
        "Drives return_value / returned_units / Net deduction."
    )
    is_stock_adjustment: str = (
        "Subtype of is_return_line that is not a confirmed return "
        "(typically Price <= 0). Drives adjustment_units only."
    )
    return_line_class: str = (
        "String label on return lines: 'confirmed_return' | 'stock_adjustment'; "
        "empty otherwise."
    )
    line_revenue: str = "Quantity * Price (signed)."

    # Counts
    distinct_invoice_count: str = (
        "nunique(Invoice) within the analysis filter. Helper flags are "
        "grouping-key-specific; non-additive Grand Totals are hidden on those pivots."
    )
    distinct_customer_count: str = (
        "nunique(Customer ID) excluding missing IDs. Same Grand Total caveat."
    )
    aov: str = "revenue_for_scope / distinct_invoice_count_for_scope (never average of row AOVs)."
    return_rate_units: str = (
        "Assignment-facing return_line_rate = "
        "(returned_units + adjustment_units) / units_sold "
        f"(eligible when units_sold >= {RETURNS_MIN_GROSS_UNITS}). "
        "Also report subtype rates: confirmed return_rate and adjustment_rate. "
        "Top-N ranks by return_line_rate (see returns_top_n)."
    )

    # Rankings (single source of truth for Summary / Pivot publication / email)
    product_top_n: int = 10
    product_rank_by: str = "gross_revenue (exclude NON_PRODUCT_STOCK_CODES from ranking)."
    country_top_n: int = 5
    country_filter: str = f"Exclude {HOME_COUNTRY}."
    customer_top_n: int = 20
    returns_top_n: int = 10
    missing_customer_country_top_n: int = 5
    customer_missing_id: str = (
        "Keep anonymous rows in overall revenue; exclude from Customer Top ranking "
        "and document revenue share + country mix of missing Customer ID."
    )

    # Time
    complete_month_rule: str = (
        "A month is complete when the dataset span covers that month's first calendar day "
        "through its last calendar day (global min/max InvoiceDate)."
    )
    sales_timing: str = (
        "peak_weekday / peak_hour are separate weekday and hour totals — not the single "
        "busiest day-hour cell (that cell is reported separately as peak_day_hour_cell)."
    )


CONTRACT = MetricsContract()


def describe_contract() -> str:
    """Human-readable contract for README / Summary sheet."""
    c = CONTRACT
    lines = [
        f"Metrics contract version: {c.version}",
        f"Gross Revenue: {c.gross_revenue}",
        f"Return Value: {c.return_value}",
        f"Stock adjustments: {c.adjustment_note}",
        f"Cancellation Value: {c.cancellation_value}",
        f"Net Revenue: {c.net_revenue}",
        f"Email Total Revenue: {c.email_total_revenue}",
        f"Cancellation: {c.is_cancellation}",
        f"Return indicator (assignment): {c.is_return_line}",
        f"Confirmed return subtype: {c.is_confirmed_return}",
        f"Stock adjustment subtype: {c.is_stock_adjustment}",
        f"AOV: {c.aov}",
        f"Return Rate: {c.return_rate_units}",
        f"Product ranking: {c.product_rank_by}; top {c.product_top_n}",
        f"Country: {c.country_filter}; top {c.country_top_n}",
        f"Customer top {c.customer_top_n}; Returns top {c.returns_top_n}",
        f"Complete months: {c.complete_month_rule}",
    ]
    return "\n".join(lines)
