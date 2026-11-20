#!/usr/bin/env python3
"""
Phase 0 criterion D: load Online Retail II, audit overlap/duplicates, row-limit check.

Does not silently drop to a sales-only subset. If cleaned rows exceed Excel's limit,
report FAIL loudly.
"""

from __future__ import annotations

import hashlib
import json
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import AUDIT_DIR, DATA_PROCESSED, DATA_RAW, INPUT_XLSX_NAME, POC_DIR, ensure_dirs  # noqa: E402
from src.metrics import CONTRACT, NON_PRODUCT_STOCK_CODES  # noqa: E402

EXCEL_MAX_DATA_ROWS = 1_048_575  # excluding header


@dataclass
class CheckResult:
    """One PASS/FAIL/SKIP check row in the Phase 0 data PoC report."""

    id: str
    name: str
    status: str
    detail: str = ""


@dataclass
class Report:
    """Mutable container for Phase 0 data PoC check results."""

    started_at: str
    finished_at: str | None = None
    checks: list[CheckResult] = field(default_factory=list)

    def add(self, check_id: str, name: str, status: str, detail: str = "") -> None:
        """Append a named check result."""
        self.checks.append(CheckResult(check_id, name, status, detail))


def find_input_workbook() -> Path:
    """Locate online_retail_II.xlsx (or *retail*.xlsx) under data/raw."""
    direct = DATA_RAW / INPUT_XLSX_NAME
    if direct.exists():
        return direct
    matches = list(DATA_RAW.glob("**/online_retail_II.xlsx")) + list(
        DATA_RAW.glob("**/*retail*.xlsx")
    )
    if matches:
        return matches[0]
    raise FileNotFoundError(
        f"Place {INPUT_XLSX_NAME} under {DATA_RAW} (UCI Online Retail II)."
    )


def file_sha256(path: Path) -> str:
    """SHA-256 hex digest of the input workbook."""
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load_combined(path: Path) -> pd.DataFrame:
    """Load all sheets and tag each row with source_sheet."""
    xl = pd.ExcelFile(path)
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(
            path,
            sheet_name=sheet,
            dtype={
                "Invoice": "string",
                "StockCode": "string",
                "Description": "string",
                "Country": "string",
            },
        )
        df["source_sheet"] = sheet
        df["source_row"] = range(2, len(df) + 2)  # Excel-like 1-index + header
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)
    combined["load_row_id"] = range(1, len(combined) + 1)
    return combined, xl.sheet_names


def classify_duplicates(df: pd.DataFrame) -> dict:
    """
    Conservative rules:
    - cross-sheet identical business key rows: flag as cross_sheet_overlap (candidate import overlap)
    - within-sheet exact duplicates: flag as exact_dup_within_sheet (do NOT auto-delete)
    - unexplained exact dups kept as kept_as_legitimate_or_suspect
    """
    key_cols = [
        "Invoice",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "Price",
        "Customer ID",
        "Country",
    ]
    work = df.copy()
    for c in key_cols:
        if c not in work.columns:
            raise KeyError(c)

    # Normalize for comparison without mutating originals beyond audit helpers
    work["_key"] = work[key_cols].astype("string").fillna("<NA>").agg("||".join, axis=1)

    within = (
        work.groupby(["source_sheet", "_key"], dropna=False)
        .size()
        .reset_index(name="n")
    )
    within_dup_keys = within.loc[within["n"] > 1, ["source_sheet", "_key"]]
    within_mask = work.set_index(["source_sheet", "_key"]).index.isin(
        within_dup_keys.set_index(["source_sheet", "_key"]).index
    )

    # Same business key appearing on more than one sheet
    sheet_per_key = work.groupby("_key")["source_sheet"].nunique()
    cross_keys = set(sheet_per_key[sheet_per_key > 1].index)
    cross_mask = work["_key"].isin(cross_keys)

    audit = pd.DataFrame(
        {
            "load_row_id": work["load_row_id"],
            "source_sheet": work["source_sheet"],
            "source_row": work["source_row"],
            "Invoice": work["Invoice"],
            "StockCode": work["StockCode"],
            "exact_dup_within_sheet": within_mask,
            "cross_sheet_overlap": cross_mask,
        }
    )
    # Conservative: do not auto-delete either class in Clean_Data for PoC.
    # Only report counts; full pipeline may delete cross-sheet import clones with documented evidence.
    return {
        "audit": audit,
        "within_dup_rows": int(within_mask.sum()),
        "cross_sheet_rows": int(cross_mask.sum()),
        "cross_sheet_keys": int(len(cross_keys)),
        "within_dup_groups": int(len(within_dup_keys)),
    }


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    """Lightweight PoC enrich (flags/revenue/time); production path uses src.data_processing.enrich."""
    out = df.copy()
    out["Invoice"] = out["Invoice"].astype("string")
    out["StockCode"] = out["StockCode"].astype("string").str.strip()
    out["is_cancellation"] = out["Invoice"].fillna("").str.upper().str.startswith("C")
    out["Quantity"] = pd.to_numeric(out["Quantity"], errors="coerce")
    out["Price"] = pd.to_numeric(out["Price"], errors="coerce")
    out["line_revenue"] = out["Quantity"] * out["Price"]
    out["is_return"] = (out["Quantity"] < 0) & (~out["is_cancellation"])
    out["gross_units"] = out["Quantity"].where(
        (out["Quantity"] > 0) & (~out["is_cancellation"]), 0
    )
    out["returned_units"] = (-out["Quantity"]).where(out["is_return"], 0)
    out["gross_revenue"] = out["line_revenue"].where(
        (out["Quantity"] > 0) & (~out["is_cancellation"]), 0.0
    )
    out["return_value"] = (-out["line_revenue"]).where(out["is_return"], 0.0)
    out["InvoiceDate"] = pd.to_datetime(out["InvoiceDate"], errors="coerce")
    out["year_month"] = out["InvoiceDate"].dt.strftime("%Y-%m")
    out["day_of_week"] = out["InvoiceDate"].dt.day_name()
    out["hour"] = out["InvoiceDate"].dt.hour
    out["is_non_product"] = out["StockCode"].str.upper().isin(
        {s.upper() for s in NON_PRODUCT_STOCK_CODES}
    )
    return out


def main() -> int:
    """Run Phase 0 data audit (criterion D) and write POC report/parquet artifacts."""
    ensure_dirs()
    report = Report(started_at=datetime.now().isoformat(timespec="seconds"))

    try:
        path = find_input_workbook()
    except FileNotFoundError as exc:
        report.add("D-input", "Locate online_retail_II.xlsx", "FAIL", str(exc))
        _write(report)
        return 1

    sha = file_sha256(path)
    report.add(
        "D-input",
        "Locate online_retail_II.xlsx",
        "PASS",
        f"{path} size={path.stat().st_size} sha256={sha}",
    )

    combined, sheets = load_combined(path)
    report.add(
        "D-load",
        "Load and combine both worksheets",
        "PASS",
        f"sheets={sheets}; rows={len(combined)}",
    )

    # Date span / sheet boundary
    dates = pd.to_datetime(combined["InvoiceDate"], errors="coerce")
    by_sheet = (
        combined.assign(_d=dates)
        .groupby("source_sheet")["_d"]
        .agg(["min", "max", "count"])
        .to_dict(orient="index")
    )
    report.add(
        "D-dates",
        "Per-sheet date ranges",
        "PASS",
        json.dumps({k: {kk: str(vv) for kk, vv in v.items()} for k, v in by_sheet.items()}),
    )

    dup = classify_duplicates(combined)
    audit_path = AUDIT_DIR / "duplicate_flags.parquet"
    dup["audit"].to_parquet(audit_path, index=False)
    report.add(
        "D-dup-audit",
        "Duplicate / overlap audit (no silent deletes)",
        "PASS",
        (
            f"within_dup_rows={dup['within_dup_rows']} groups={dup['within_dup_groups']}; "
            f"cross_sheet_rows={dup['cross_sheet_rows']} keys={dup['cross_sheet_keys']}; "
            f"audit={audit_path.name}. Conservative policy: keep all rows in Clean_Data candidate."
        ),
    )

    enriched = enrich(combined)
    # Clean candidate = all rows with required types; invalid numeric flagged not dropped here
    invalid_qty_price = enriched["Quantity"].isna() | enriched["Price"].isna()
    missing_invoice = enriched["Invoice"].isna() | (enriched["Invoice"].str.len() == 0)
    report.add(
        "D-quality-flags",
        "Flag invalid/missing without dropping for PoC",
        "PASS",
        (
            f"invalid_qty_or_price={int(invalid_qty_price.sum())}; "
            f"missing_invoice={int(missing_invoice.sum())}; "
            f"cancellations={int(enriched['is_cancellation'].sum())}; "
            f"returns={int(enriched['is_return'].sum())}; "
            f"missing_customer_id={int(enriched['Customer ID'].isna().sum())}"
        ),
    )

    clean_rows = len(enriched)
    fits = clean_rows <= EXCEL_MAX_DATA_ROWS
    report.add(
        "D-excel-row-limit",
        "Clean_Data row count vs Excel limit",
        "PASS" if fits else "FAIL",
        (
            f"clean_candidate_rows={clean_rows}; excel_max={EXCEL_MAX_DATA_ROWS}; "
            f"fits={fits}. No sales-only subset fallback."
        ),
    )

    # Persist processed sample artifacts
    parquet_path = DATA_PROCESSED / "transactions_enriched.parquet"
    enriched.to_parquet(parquet_path, index=False)
    summary = {
        "input": str(path),
        "sha256": sha,
        "sheets": sheets,
        "raw_rows": int(len(combined)),
        "clean_candidate_rows": clean_rows,
        "excel_max_data_rows": EXCEL_MAX_DATA_ROWS,
        "fits_excel": fits,
        "metrics_contract_version": CONTRACT.version,
        "duplicate_audit": {
            "within_dup_rows": dup["within_dup_rows"],
            "within_dup_groups": dup["within_dup_groups"],
            "cross_sheet_rows": dup["cross_sheet_rows"],
            "cross_sheet_keys": dup["cross_sheet_keys"],
        },
        "date_min": str(dates.min()),
        "date_max": str(dates.max()),
    }
    (POC_DIR / "phase0_data_summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )
    report.add(
        "D-persist",
        "Persist enriched parquet + summary",
        "PASS",
        f"{parquet_path} rows={clean_rows}",
    )

    report.finished_at = datetime.now().isoformat(timespec="seconds")
    _write(report)
    failed = sum(1 for c in report.checks if c.status == "FAIL")
    return 1 if failed else 0


def _write(report: Report) -> None:
    """Persist phase0_data_report.json under POC_DIR and print a console summary."""
    out = {
        "started_at": report.started_at,
        "finished_at": report.finished_at or datetime.now().isoformat(timespec="seconds"),
        "checks": [asdict(c) for c in report.checks],
    }
    path = POC_DIR / "phase0_data_report.json"
    path.write_text(json.dumps(out, indent=2), encoding="utf-8")
    print("=== Phase 0 Data PoC Results ===")
    for c in report.checks:
        print(f"[{c.status}] {c.id}: {c.name}")
        if c.detail:
            print(f"         {c.detail[:700]}")
    print(f"\nReport: {path}")


if __name__ == "__main__":
    raise SystemExit(main())
