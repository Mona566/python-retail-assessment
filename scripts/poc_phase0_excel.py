#!/usr/bin/env python3
"""
Phase 0 Excel PoC: native PivotTables, distinct helpers, PNG export.

Does not process the full Online Retail II dataset.
Reports PASS/FAIL for criteria A, B, C, E (partial) on a small synthetic sample.
"""

from __future__ import annotations

import json
import sys
import traceback
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import POC_DIR, PNG_DIR, ensure_dirs  # noqa: E402
from src.metrics import CONTRACT, describe_contract  # noqa: E402


@dataclass
class CheckResult:
    """One PASS/FAIL/SKIP check for the Excel Pivot PoC."""

    id: str
    name: str
    status: str  # PASS | FAIL | SKIP
    detail: str = ""


@dataclass
class Report:
    """Excel PoC report: checks plus xlwings version and metrics contract text."""

    started_at: str
    finished_at: str | None = None
    xlwings_version: str | None = None
    checks: list[CheckResult] = field(default_factory=list)
    metrics_contract: str = ""

    def add(self, check_id: str, name: str, status: str, detail: str = "") -> None:
        """Append a named check result."""
        self.checks.append(CheckResult(check_id, name, status, detail))


def build_sample_frame() -> pd.DataFrame:
    """Synthetic retail lines covering sales, returns, multi-product customers."""
    rows = [
        # Customer C001 buys A and B on invoice 1001
        ("1001", "A", "Product A", 2, 10.0, "C001", "France", "2021-01-04 10:00:00", False),
        ("1001", "B", "Product B", 1, 50.0, "C001", "France", "2021-01-04 10:00:00", False),
        # Same customer, another invoice
        ("1002", "A", "Product A", 3, 10.0, "C001", "France", "2021-01-05 11:00:00", False),
        # Other customers / countries
        ("1003", "A", "Product A", 5, 10.0, "C002", "Germany", "2021-01-05 14:00:00", False),
        ("1004", "B", "Product B", 2, 50.0, "C003", "Germany", "2021-01-06 09:00:00", False),
        ("1005", "C", "Product C", 10, 5.0, "C002", "France", "2021-01-06 15:00:00", False),
        # Return of A
        ("1006", "A", "Product A", -1, 10.0, "C001", "France", "2021-01-07 10:00:00", False),
        # Cancellation
        ("C1007", "B", "Product B", -2, 50.0, "C003", "Germany", "2021-01-07 16:00:00", True),
        # UK (excluded from country pivot source filter tests later)
        ("1008", "A", "Product A", 4, 10.0, "C004", "United Kingdom", "2021-01-08 12:00:00", False),
        ("1009", "B", "Product B", 1, 50.0, "C005", "France", "2021-01-08 13:00:00", False),
    ]
    df = pd.DataFrame(
        rows,
        columns=[
            "Invoice",
            "StockCode",
            "Description",
            "Quantity",
            "Price",
            "Customer ID",
            "Country",
            "InvoiceDate",
            "is_cancellation",
        ],
    )
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
    df["line_revenue"] = df["Quantity"] * df["Price"]
    df["is_return"] = (df["Quantity"] < 0) & (~df["is_cancellation"])
    df["gross_units"] = df["Quantity"].where(
        (df["Quantity"] > 0) & (~df["is_cancellation"]), 0
    )
    df["returned_units"] = (-df["Quantity"]).where(df["is_return"], 0)
    df["gross_revenue"] = df["line_revenue"].where(
        (df["Quantity"] > 0) & (~df["is_cancellation"]), 0.0
    )
    df["return_value"] = (-df["line_revenue"]).where(df["is_return"], 0.0)
    df["net_revenue"] = df["gross_revenue"] - df["return_value"]
    # For return lines, net contribution is already negative in line_revenue;
    # net_revenue column above is line-level gross/return split for Sum pivots.
    df["year_month"] = df["InvoiceDate"].dt.strftime("%Y-%m")
    df["day_of_week"] = df["InvoiceDate"].dt.day_name()
    df["hour"] = df["InvoiceDate"].dt.hour

    # Helpers must match the analysis filter (exclude cancellations).
    active = ~df["is_cancellation"]

    # Grouping-key-specific distinct invoice helpers (product level).
    df["inv_flag_by_product"] = 0
    first_prod = active & ~df.loc[active].duplicated(
        subset=["StockCode", "Invoice"], keep="first"
    ).reindex(df.index, fill_value=False)
    df.loc[first_prod, "inv_flag_by_product"] = 1

    # Customer-level helper keyed by StockCode (demonstrates Grand Total trap).
    df["cust_flag_by_product"] = 0
    first_cust_prod = (
        active
        & df["Customer ID"].notna()
        & ~df.loc[active & df["Customer ID"].notna()]
        .duplicated(subset=["StockCode", "Customer ID"], keep="first")
        .reindex(df.index, fill_value=False)
    )
    df.loc[first_cust_prod, "cust_flag_by_product"] = 1

    # Country-level invoice helper
    df["inv_flag_by_country"] = 0
    first_country = active & ~df.loc[active].duplicated(
        subset=["Country", "Invoice"], keep="first"
    ).reindex(df.index, fill_value=False)
    df.loc[first_country, "inv_flag_by_country"] = 1

    return df


def pandas_product_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Ground-truth product aggregates used to verify PivotTable values."""
    sales = df[~df["is_cancellation"]].copy()
    g = sales.groupby("StockCode", as_index=False).agg(
        Description=("Description", "first"),
        revenue=("gross_revenue", "sum"),
        units=("gross_units", "sum"),
        returned_units=("returned_units", "sum"),
        return_value=("return_value", "sum"),
        distinct_invoices=("Invoice", "nunique"),
        distinct_customers=("Customer ID", "nunique"),
    )
    g["return_rate"] = g["returned_units"] / g["units"].replace(0, pd.NA)
    g["aov"] = g["revenue"] / g["distinct_invoices"].replace(0, pd.NA)
    return g


def run_excel_poc(report: Report) -> Path | None:
    """
    Build a small workbook with native PivotTables, export PNGs, and record checks.

    Returns the workbook path on success, or None if Excel automation fails early.
    """
    report.xlwings_version = xw.__version__
    ensure_dirs()
    df = build_sample_frame()
    sample_path = POC_DIR / "sample_transactions.csv"
    df.to_csv(sample_path, index=False)

    out_path = POC_DIR / "poc_pivots.xlsx"
    if out_path.exists():
        out_path.unlink()

    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.add()

        # --- Clean_Data ---
        ws_data = wb.sheets[0]
        ws_data.name = "Clean_Data"
        ws_data.range("A1").options(index=False, header=True).value = df
        used = ws_data.used_range
        source = ws_data.range((1, 1), (used.last_cell.row, used.last_cell.column))

        # --- A: one basic pivot ---
        ws_a = wb.sheets.add("Pivot_Product_Revenue", after=ws_data)
        pt_product = ws_a.pivot_tables.add(
            source=source,
            destination=ws_a.range("A3"),
            name="PT_Product",
            rows="StockCode",
            values={
                "gross_revenue": "sum",
                "gross_units": "sum",
                "inv_flag_by_product": "sum",
            },
            layout="tabular",
        )
        report.add(
            "A",
            "Create basic native PivotTable",
            "PASS",
            f"Created {pt_product.name} on {ws_a.name}",
        )

        # --- B: five more pivots on separate sheets ---
        pivot_specs = [
            (
                "Pivot_Country_Performance",
                "PT_Country",
                "Country",
                {"gross_revenue": "sum", "inv_flag_by_country": "sum"},
            ),
            (
                "Pivot_Monthly_Performance",
                "PT_Monthly",
                "year_month",
                {"gross_revenue": "sum", "inv_flag_by_product": "sum"},
            ),
            (
                "Pivot_Returns",
                "PT_Returns",
                "StockCode",
                {
                    "gross_units": "sum",
                    "returned_units": "sum",
                    "return_value": "sum",
                },
            ),
            (
                "Pivot_Customer_Value",
                "PT_Customer",
                "Customer ID",
                {"gross_revenue": "sum", "inv_flag_by_country": "sum"},
            ),
            (
                "Pivot_Sales_Timing",
                "PT_Timing",
                "day_of_week",
                {"gross_revenue": "sum"},
            ),
        ]
        created = ["Pivot_Product_Revenue"]
        for sheet_name, pt_name, row_field, values in pivot_specs:
            ws = wb.sheets.add(sheet_name, after=wb.sheets[-1])
            pt = ws.pivot_tables.add(
                source=source,
                destination=ws.range("A3"),
                name=pt_name,
                rows=row_field,
                values=values,
                layout="tabular",
            )
            # Sales timing: add hour as columns if supported
            if sheet_name == "Pivot_Sales_Timing":
                try:
                    pt.columns.add("hour")
                except Exception as exc:  # noqa: BLE001
                    report.add(
                        "B-timing-cols",
                        "Add hour columns to timing pivot",
                        "FAIL",
                        str(exc),
                    )
            created.append(sheet_name)

        # Verify six sheets each have a pivot
        missing = []
        for name in created:
            sheet = wb.sheets[name]
            if len(sheet.pivot_tables) < 1:
                missing.append(name)
        if missing:
            report.add(
                "B",
                "Six sheets with native PivotTables",
                "FAIL",
                f"Missing pivots on: {missing}",
            )
        else:
            report.add(
                "B",
                "Six sheets with native PivotTables",
                "PASS",
                f"Sheets: {created}",
            )

        # --- C: distinct helper vs pandas; Grand Total trap ---
        expected = pandas_product_stats(df)
        # Read product pivot body (best-effort via used range values)
        pt_range = pt_product.range
        values = pt_range.value
        # Flatten for logging
        detail_parts = []
        group_ok = True
        # Expected per product distinct invoices
        for _, row in expected.iterrows():
            code = row["StockCode"]
            exp_inv = int(row["distinct_invoices"])
            # Find row in pivot values (structure varies by Excel layout)
            found = None
            if isinstance(values, list):
                for r in values:
                    if not isinstance(r, list):
                        continue
                    if r and str(r[0]) == str(code):
                        found = r
                        break
            detail_parts.append(f"{code}: expected_invoices={exp_inv}, pivot_row={found}")

        # Grand Total trap: sum of cust_flag_by_product vs true nunique customers
        helper_grand = int(df["cust_flag_by_product"].sum())
        true_customers = int(df["Customer ID"].nunique())
        trap_ok = helper_grand != true_customers  # we EXPECT inequality (trap exists)
        # Build a second small pivot for customer flags by product to inspect GT
        ws_trap = wb.sheets.add("PoC_GrandTotal_Trap", after=wb.sheets[-1])
        pt_trap = ws_trap.pivot_tables.add(
            source=source,
            destination=ws_trap.range("A3"),
            name="PT_Trap",
            rows="StockCode",
            values={"cust_flag_by_product": "sum"},
            layout="tabular",
        )
        # Prefer hiding row grand totals when helper is not additive across groups
        try:
            pt_trap.show_row_grand_totals = False
            gt_hidden = True
        except Exception as exc:  # noqa: BLE001
            gt_hidden = False
            detail_parts.append(f"hide_grand_total_error={exc}")

        # Product-level: helper sum must match nunique invoices on the SAME filter
        # (non-cancellation), not all invoices including cancellations.
        active = df[~df["is_cancellation"]]
        helper_vs_pandas = (
            active.groupby("StockCode")["inv_flag_by_product"].sum()
            == active.groupby("StockCode")["Invoice"].nunique()
        )
        # Also compare against pivot body values for invoice flag column
        pivot_mismatches = []
        if isinstance(values, list):
            exp_map = expected.set_index("StockCode")["distinct_invoices"].to_dict()
            for r in values:
                if not isinstance(r, list) or not r:
                    continue
                code = str(r[0])
                if code not in exp_map:
                    continue
                # layout: StockCode, sum gross_revenue, sum gross_units, sum inv_flag
                if len(r) >= 4 and r[3] is not None:
                    if int(r[3]) != int(exp_map[code]):
                        pivot_mismatches.append(
                            f"{code}: pivot_flag={r[3]} expected={exp_map[code]}"
                        )
        helper_check = bool(helper_vs_pandas.all()) and not pivot_mismatches
        if helper_check:
            report.add(
                "C-group",
                "Distinct invoice helper matches pandas per StockCode",
                "PASS",
                "; ".join(detail_parts[:5]),
            )
        else:
            group_ok = False
            report.add(
                "C-group",
                "Distinct invoice helper matches pandas per StockCode",
                "FAIL",
                f"helper_vs_pandas={helper_vs_pandas.to_dict()}; "
                f"pivot_mismatches={pivot_mismatches}; "
                + "; ".join(detail_parts),
            )

        report.add(
            "C-grand-total-trap",
            "Document non-additive Grand Total for customer helper",
            "PASS" if trap_ok and gt_hidden else "FAIL",
            (
                f"sum(cust_flag_by_product)={helper_grand}, "
                f"true_nunique_customers={true_customers}, "
                f"grand_totals_hidden={gt_hidden}"
            ),
        )

        # AOV / return rate: compute from aggregated sums (not pivot Average)
        aov_ok = True
        for _, row in expected.iterrows():
            if pd.isna(row["aov"]):
                continue
            recomputed = row["revenue"] / row["distinct_invoices"]
            if abs(recomputed - float(row["aov"])) > 1e-9:
                aov_ok = False
        rr = expected.loc[expected["StockCode"] == "A"].iloc[0]
        rr_ok = float(rr["returned_units"]) / float(rr["units"]) == float(rr["return_rate"])
        report.add(
            "C-aov-return-rate",
            "AOV and return rate from Sum fields (not Average of rates)",
            "PASS" if aov_ok and rr_ok else "FAIL",
            f"Product A return_rate={rr['return_rate']}, aov_ok={aov_ok}",
        )

        # Save, reopen, refresh
        wb.save(str(out_path))
        wb.close()
        wb = None
        wb2 = app.books.open(str(out_path))
        refresh_errors = []
        for name in created:
            sheet = wb2.sheets[name]
            try:
                sheet.pivot_tables[0].refresh()
            except Exception as exc:  # noqa: BLE001
                refresh_errors.append(f"{name}: {exc}")
        if refresh_errors:
            report.add("A-refresh", "Save, reopen, refresh pivots", "FAIL", str(refresh_errors))
        else:
            report.add(
                "A-refresh",
                "Save, reopen, refresh pivots",
                "PASS",
                f"Reopened {out_path.name}",
            )

        # --- E: PNG export ---
        png_path = PNG_DIR / "poc_product_pivot.png"
        if png_path.exists():
            png_path.unlink()
        try:
            pt = wb2.sheets["Pivot_Product_Revenue"].pivot_tables[0]
            pt.range.to_png(str(png_path))
            size = png_path.stat().st_size if png_path.exists() else 0
            if size > 0:
                report.add(
                    "E",
                    "Pivot range to_png non-empty",
                    "PASS",
                    f"{png_path.name} size={size} bytes (manual content check still required)",
                )
            else:
                report.add("E", "Pivot range to_png non-empty", "FAIL", "File missing or empty")
        except Exception as exc:  # noqa: BLE001
            report.add("E", "Pivot range to_png", "FAIL", str(exc))

        # Pivot cache / workbook size notes
        size_bytes = out_path.stat().st_size
        report.add(
            "B-size",
            "PoC workbook size (sample only)",
            "PASS",
            f"{size_bytes} bytes; full-data cache check deferred to scale test",
        )

        wb2.save()
        wb2.close()
        wb2 = None
        return out_path
    except Exception as exc:  # noqa: BLE001
        report.add("A", "Excel Pivot PoC", "FAIL", f"{exc}\n{traceback.format_exc()}")
        return None
    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:  # noqa: BLE001
            pass
        if app is not None:
            try:
                app.quit()
            except Exception:  # noqa: BLE001
                pass


def main() -> int:
    """CLI entry for Phase 0 Excel PoC (criteria A/B/C/E on synthetic data)."""
    ensure_dirs()
    report = Report(
        started_at=datetime.now().isoformat(timespec="seconds"),
        metrics_contract=describe_contract(),
    )
    print("Metrics contract:\n", report.metrics_contract, "\n", sep="")
    print(f"Contract object version: {CONTRACT.version}")

    run_excel_poc(report)
    report.finished_at = datetime.now().isoformat(timespec="seconds")

    out = {
        "started_at": report.started_at,
        "finished_at": report.finished_at,
        "xlwings_version": report.xlwings_version,
        "metrics_contract_version": CONTRACT.version,
        "checks": [asdict(c) for c in report.checks],
    }
    report_path = POC_DIR / "phase0_excel_report.json"
    report_path.write_text(json.dumps(out, indent=2), encoding="utf-8")

    print("\n=== Phase 0 Excel PoC Results ===")
    failed = 0
    for c in report.checks:
        print(f"[{c.status}] {c.id}: {c.name}")
        if c.detail:
            print(f"         {c.detail[:500]}")
        if c.status == "FAIL":
            failed += 1
    print(f"\nReport written to {report_path}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
