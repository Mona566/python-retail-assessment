"""Pivot specs, Excel column renames, and value number formats for Retail_Analysis.xlsx."""

from __future__ import annotations

from src.metrics import CONTRACT, HOME_COUNTRY

# Rename technical helper columns → readable Excel headers BEFORE building pivots.
# Native PivotTables then show "Sum of Revenue (GBP)" instead of "Sum of inv_flag_…".
CLEAN_DATA_EXCEL_RENAME: dict[str, str] = {
    "product_label": "Product",
    "Country_NonUK": "Country (excl. UK)",
    "year_month": "Year-month",
    "weekday_label": "Day of week",
    "hour": "Hour",
    "gross_revenue": "Revenue (GBP)",
    "gross_revenue_product": "Product revenue (GBP)",
    "gross_units_product": "Product units sold",
    "inv_flag_by_product_only": "Distinct invoices",
    "gross_revenue_non_uk": "Non-UK revenue (GBP)",
    "inv_flag_by_country_non_uk": "Non-UK distinct invoices",
    "cust_flag_by_country_non_uk": "Non-UK known customers",
    "inv_flag_by_month": "Monthly distinct invoices",
    "cust_flag_by_month": "Monthly known customers",
    "gross_revenue_known_customer": "Known-customer revenue (GBP)",
    "inv_flag_by_customer": "Customer distinct invoices",
    "line_revenue": "Line revenue (GBP)",
    "return_value": "Confirmed return value (GBP)",
    "gross_units": "Gross units",
    "returned_units": "Confirmed returned units",
    "adjustment_units": "Stock adjustment units",
    "cancellation_value_line": "Cancellation value (GBP)",
    "cancelled_units": "Cancelled units",
}

RETURNS_SOURCE_EXCEL_RENAME: dict[str, str] = {
    "product_label": "Product",
    "units_sold": "Units sold",
    "return_line_units": "Return-line units",
    "returned_units": "Confirmed returned units",
    "adjustment_units": "Stock adjustment units",
    "return_value": "Confirmed return value (GBP)",
    "return_line_rate": "Negative-quantity line rate",
    "return_rate": "Confirmed return rate",
    "adjustment_rate": "Adjustment rate",
    "cancelled_units": "Cancelled units",
    "cancellation_value": "Cancellation value (GBP)",
}

# Excel header → number format (applied on Clean_Data / Returns_Source + Pivot data fields)
PIVOT_VALUE_STYLE: dict[str, tuple[str, str]] = {
    "Average order value (GBP)": ("Average order value (GBP)", "£#,##0.00"),
    "Revenue rank": ("Revenue rank", "#,##0"),
    "Product revenue (GBP)": ("Product revenue (GBP)", "£#,##0.00"),
    "Product units sold": ("Product units sold", "#,##0"),
    "Distinct invoices": ("Distinct invoices", "#,##0"),
    "Non-UK revenue (GBP)": ("Non-UK revenue (GBP)", "£#,##0.00"),
    "Non-UK distinct invoices": ("Non-UK distinct invoices", "#,##0"),
    "Non-UK known customers": ("Non-UK known customers", "#,##0"),
    "Revenue (GBP)": ("Revenue (GBP)", "£#,##0.00"),
    "Monthly distinct invoices": ("Monthly distinct invoices", "#,##0"),
    "Monthly known customers": ("Monthly known customers", "#,##0"),
    "Units sold": ("Units sold", "#,##0"),
    "Return-line units": ("Return-line units", "#,##0"),
    "Confirmed returned units": ("Confirmed returned units", "#,##0"),
    "Stock adjustment units": ("Stock adjustment units", "#,##0"),
    "Confirmed return value (GBP)": ("Confirmed return value (GBP)", "£#,##0.00"),
    "Negative-quantity line rate": ("Negative-quantity line rate", "0.00%"),
    "Confirmed return rate": ("Confirmed return rate", "0.00%"),
    "Adjustment rate": ("Adjustment rate", "0.00%"),
    "Cancelled units": ("Cancelled units", "#,##0"),
    "Cancellation value (GBP)": ("Cancellation value (GBP)", "£#,##0.00"),
    "Known-customer revenue (GBP)": ("Known-customer revenue (GBP)", "£#,##0.00"),
    "Customer distinct invoices": ("Customer distinct invoices", "#,##0"),
}

PIVOT_SPECS = [
    {
        "sheet": "Pivot_Product_Revenue",
        "name": "PT_Product_Revenue",
        "source_sheet": "Product_Source",
        "rows": "Product",
        "values": {
            "Product revenue (GBP)": "sum",
            "Product units sold": "sum",
            "Distinct invoices": "sum",
        },
        "hide_bottom_grand_total": True,
        "note": (
            f"Rows are StockCode — Description. Top {CONTRACT.product_top_n} by revenue. "
            "Invoice counts are computed per product in Python."
        ),
    },
    {
        "sheet": "Pivot_Country_Performance",
        "name": "PT_Country_Performance",
        "source_sheet": "Country_Source",
        "rows": "Country (excl. UK)",
        "values": {
            "Non-UK revenue (GBP)": "sum",
            "Non-UK distinct invoices": "sum",
            "Non-UK known customers": "sum",
            "Average order value (GBP)": "average",
        },
        "hide_bottom_grand_total": True,
        "note": (
            f"{HOME_COUNTRY} excluded. Top {CONTRACT.country_top_n} by revenue. "
            "AOV = revenue / distinct invoices. "
            "Customer count = known IDs only; anonymous revenue still in country totals "
            "(see Summary country missing-ID note)."
        ),
    },
    {
        "sheet": "Pivot_Monthly_Performance",
        "name": "PT_Monthly_Performance",
        "source_sheet": "Monthly_Source",
        "rows": "Year-month",
        "values": {
            "Revenue (GBP)": "sum",
            "Monthly distinct invoices": "sum",
            "Monthly known customers": "sum",
        },
        "hide_bottom_grand_total": True,
        "note": (
            "Native Pivot on the Python monthly rollup. Complete vs incomplete months are labeled on "
            "Summary and listed on Monthly_Source (edge months excluded from best/worst)."
        ),
    },
    {
        "sheet": "Pivot_Returns",
        "name": "PT_Returns",
        "source_sheet": "Returns_Source",
        "rows": "Product",
        "values": {
            "Units sold": "sum",
            "Return-line units": "sum",
            "Confirmed returned units": "sum",
            "Stock adjustment units": "sum",
            "Confirmed return value (GBP)": "sum",
            "Negative-quantity line rate": "sum",
            "Confirmed return rate": "sum",
            "Adjustment rate": "sum",
            "Cancelled units": "sum",
            "Cancellation value (GBP)": "sum",
        },
        # Rates are pre-computed per product; a bottom Grand Total would sum rates
        # (meaningless). Hide it; row-level rates remain valid.
        "hide_bottom_grand_total": True,
        "note": (
            "Native PivotTable sourced from Returns_Source (Python-pre-aggregated eligible "
            "products with units_sold >= threshold) — still a live Excel Pivot, not a pasted "
            f"table. Top {CONTRACT.returns_top_n} by negative-quantity line rate "
            "(includes stock adjustments); confirmed vs stock-adjustment subtypes shown. "
            "Bottom Grand Total hidden because rate columns are not additive."
        ),
    },
    {
        "sheet": "Pivot_Customer_Value",
        "name": "PT_Customer_Value",
        "source_sheet": "Customer_Source",
        "rows": "Customer ID",
        "values": {
            "Known-customer revenue (GBP)": "sum",
            "Customer distinct invoices": "sum",
            "Average order value (GBP)": "average",
            "Revenue rank": "min",
        },
        "hide_bottom_grand_total": True,
        "note": (
            f"Top {CONTRACT.customer_top_n} by revenue. "
            "Customers without an ID are excluded."
        ),
    },
    {
        "sheet": "Pivot_Sales_Timing",
        "name": "PT_Sales_Timing",
        "source_sheet": "Timing_Source",
        "rows": "Day of week",
        "columns": "Hour",
        "values": {"Revenue (GBP)": "sum"},
        "hide_bottom_grand_total": False,
        "note": "Matrix: day-of-week × hour; values = Revenue (GBP).",
    },
]
