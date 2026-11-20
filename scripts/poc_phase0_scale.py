#!/usr/bin/env python3
"""
Phase 0 criterion H: Excel Pivot stress on 100k rows, then optional fuller sample.

Usage:
  python scripts/poc_phase0_scale.py           # 100k
  python scripts/poc_phase0_scale.py --full    # all enriched rows (may take long)
"""

from __future__ import annotations

import json
import sys
import time
import traceback
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import DATA_PROCESSED, PNG_DIR, POC_DIR, ensure_dirs  # noqa: E402

PIVOT_SHEETS = [
    ("Pivot_Product_Revenue", "PT_Product", "StockCode", {
        "gross_revenue": "sum",
        "gross_units": "sum",
        "inv_flag_by_product": "sum",
    }),
    ("Pivot_Country_Performance", "PT_Country", "Country", {
        "gross_revenue": "sum",
        "inv_flag_by_country": "sum",
    }),
    ("Pivot_Monthly_Performance", "PT_Monthly", "year_month", {
        "gross_revenue": "sum",
        "inv_flag_by_product": "sum",
    }),
    ("Pivot_Returns", "PT_Returns", "StockCode", {
        "gross_units": "sum",
        "returned_units": "sum",
        "return_value": "sum",
    }),
    ("Pivot_Customer_Value", "PT_Customer", "Customer ID", {
        "gross_revenue": "sum",
        "inv_flag_by_country": "sum",
    }),
    ("Pivot_Sales_Timing", "PT_Timing", "day_of_week", {
        "gross_revenue": "sum",
    }),
]


@dataclass
class CheckResult:
    """One timing/correctness check for the scale Pivot PoC."""

    id: str
    name: str
    status: str
    detail: str = ""


@dataclass
class Report:
    """Scale PoC report container."""

    started_at: str
    finished_at: str | None = None
    checks: list[CheckResult] = field(default_factory=list)

    def add(self, check_id: str, name: str, status: str, detail: str = "") -> None:
        """Append a named check result."""
        self.checks.append(CheckResult(check_id, name, status, detail))


def prepare_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure cancellation flags and invoice helpers exist; return Pivot-ready columns."""
    out = df.copy()
    if "is_cancellation" not in out.columns:
        out["is_cancellation"] = (
            out["Invoice"].astype("string").fillna("").str.upper().str.startswith("C")
        )
    active = ~out["is_cancellation"].fillna(False)
    out["inv_flag_by_product"] = 0
    mask = active & ~out.loc[active].duplicated(
        subset=["StockCode", "Invoice"], keep="first"
    ).reindex(out.index, fill_value=False)
    out.loc[mask, "inv_flag_by_product"] = 1

    out["inv_flag_by_country"] = 0
    mask_c = active & ~out.loc[active].duplicated(
        subset=["Country", "Invoice"], keep="first"
    ).reindex(out.index, fill_value=False)
    out.loc[mask_c, "inv_flag_by_country"] = 1

    # Keep Excel-friendly subset of columns for write speed
    cols = [
        "Invoice",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "Price",
        "Customer ID",
        "Country",
        "source_sheet",
        "source_row",
        "is_cancellation",
        "is_return",
        "gross_units",
        "returned_units",
        "gross_revenue",
        "return_value",
        "year_month",
        "day_of_week",
        "hour",
        "inv_flag_by_product",
        "inv_flag_by_country",
    ]
    missing = [c for c in cols if c not in out.columns]
    if missing:
        raise KeyError(f"Missing columns for scale PoC: {missing}")
    return out[cols]


def run_scale(df: pd.DataFrame, label: str, report: Report) -> None:
    """
    Stress-test native Pivot creation/refresh/PNG on ``df`` (label e.g. 100k/full).

    Writes ``poc_scale_{label}.xlsx`` under POC_DIR and appends timing checks to report.
    """
    ensure_dirs()
    out_path = POC_DIR / f"poc_scale_{label}.xlsx"
    if out_path.exists():
        out_path.unlink()

    t0 = time.perf_counter()
    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.add()
        ws_data = wb.sheets[0]
        ws_data.name = "Clean_Data"

        t_write0 = time.perf_counter()
        ws_data.range("A1").options(index=False, header=True).value = df
        t_write = time.perf_counter() - t_write0
        used = ws_data.used_range
        source = ws_data.range((1, 1), (used.last_cell.row, used.last_cell.column))
        report.add(
            f"H-{label}-write",
            f"Write {len(df)} rows to Clean_Data",
            "PASS",
            f"seconds={t_write:.1f}; used={used.address}",
        )

        created = []
        t_pivot0 = time.perf_counter()
        for sheet_name, pt_name, row_field, values in PIVOT_SHEETS:
            ws = wb.sheets.add(sheet_name, after=wb.sheets[-1])
            pt = ws.pivot_tables.add(
                source=source,
                destination=ws.range("A3"),
                name=pt_name,
                rows=row_field,
                values=values,
                layout="tabular",
            )
            try:
                pt.show_row_grand_totals = False
            except Exception:  # noqa: BLE001
                pass
            if sheet_name == "Pivot_Sales_Timing":
                try:
                    pt.columns.add("hour")
                except Exception as exc:  # noqa: BLE001
                    report.add(
                        f"H-{label}-timing-cols",
                        "Add hour columns",
                        "FAIL",
                        str(exc),
                    )
            created.append(sheet_name)
        t_pivot = time.perf_counter() - t_pivot0

        missing = [n for n in created if len(wb.sheets[n].pivot_tables) < 1]
        report.add(
            f"H-{label}-pivots",
            "Create six native PivotTables",
            "FAIL" if missing else "PASS",
            f"seconds={t_pivot:.1f}; missing={missing or None}",
        )

        t_save0 = time.perf_counter()
        wb.save(str(out_path))
        t_save = time.perf_counter() - t_save0
        size = out_path.stat().st_size
        wb.close()
        wb = None

        wb2 = app.books.open(str(out_path))
        refresh_errs = []
        for name in created:
            try:
                wb2.sheets[name].pivot_tables[0].refresh()
            except Exception as exc:  # noqa: BLE001
                refresh_errs.append(f"{name}:{exc}")
        report.add(
            f"H-{label}-reopen-refresh",
            "Reopen and refresh all pivots",
            "FAIL" if refresh_errs else "PASS",
            f"save_seconds={t_save:.1f}; size_bytes={size}; errs={refresh_errs or None}",
        )

        # PNG for first pivot only at scale (time); full six deferred if huge
        png = PNG_DIR / f"poc_scale_{label}_product.png"
        if png.exists():
            png.unlink()
        try:
            pt = wb2.sheets["Pivot_Product_Revenue"].pivot_tables[0]
            # Hide noisy totals already; export visible range
            pt.range.to_png(str(png))
            psz = png.stat().st_size if png.exists() else 0
            report.add(
                f"H-{label}-png",
                "Export product pivot PNG",
                "PASS" if psz > 0 else "FAIL",
                f"{png.name} bytes={psz}",
            )
        except Exception as exc:  # noqa: BLE001
            report.add(f"H-{label}-png", "Export product pivot PNG", "FAIL", str(exc))

        wb2.save()
        wb2.close()
        wb2 = None
        total = time.perf_counter() - t0
        report.add(
            f"H-{label}-total",
            f"Scale run complete ({label})",
            "PASS",
            f"total_seconds={total:.1f}; workbook={out_path.name}",
        )
    except Exception as exc:  # noqa: BLE001
        report.add(
            f"H-{label}",
            f"Scale Excel PoC ({label})",
            "FAIL",
            f"{exc}\n{traceback.format_exc()}",
        )
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
    """CLI entry for criterion H: 100k (default) or --full enriched-row Pivot stress."""
    ensure_dirs()
    report = Report(started_at=datetime.now().isoformat(timespec="seconds"))
    parquet = DATA_PROCESSED / "transactions_enriched.parquet"
    if not parquet.exists():
        report.add(
            "H-prereq",
            "Need enriched parquet from poc_phase0_data.py",
            "FAIL",
            str(parquet),
        )
        _write(report)
        return 1

    df_all = pd.read_parquet(parquet)
    report.add("H-prereq", "Load enriched parquet", "PASS", f"rows={len(df_all)}")

    n = len(df_all) if "--full" in sys.argv else min(100_000, len(df_all))
    label = "full" if "--full" in sys.argv else "100k"
    sample = prepare_frame(df_all.head(n).copy())
    # For more representative 100k, sample across time if not full
    if label == "100k" and len(df_all) > 100_000:
        sample = prepare_frame(df_all.sample(n=100_000, random_state=42).copy())

    run_scale(sample, label, report)
    report.finished_at = datetime.now().isoformat(timespec="seconds")
    _write(report)
    return 1 if any(c.status == "FAIL" for c in report.checks) else 0


def _write(report: Report) -> None:
    """Persist phase0_scale_report.json and print a console summary."""
    path = POC_DIR / "phase0_scale_report.json"
    path.write_text(
        json.dumps(
            {
                "started_at": report.started_at,
                "finished_at": report.finished_at
                or datetime.now().isoformat(timespec="seconds"),
                "checks": [asdict(c) for c in report.checks],
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    print("=== Phase 0 Scale PoC Results ===")
    for c in report.checks:
        print(f"[{c.status}] {c.id}: {c.name}")
        if c.detail:
            print(f"         {c.detail[:700]}")
    print(f"\nReport: {path}")


if __name__ == "__main__":
    raise SystemExit(main())
