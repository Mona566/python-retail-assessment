#!/usr/bin/env python3
"""
Micro-benchmarks for prep hotspots (vectorization + independent I/O threads).

Records wall times with ``time.perf_counter()`` so README numbers are measured,
not guessed. Run from repo root:

  python scripts/bench_hotspots.py
  python scripts/bench_hotspots.py --rows 200000
"""

from __future__ import annotations

import argparse
import json
import platform
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.data_processing import (  # noqa: E402
    BUSINESS_KEY_COLS,
    _business_key,
    _description_mode_by_stock,
    _run_io_tasks,
)
from src.logging_utils import _peak_rss_mb  # noqa: E402


def _synth(n: int) -> pd.DataFrame:
    """Synthetic ledger large enough to expose string/groupby costs."""
    rng = pd.RangeIndex(n)
    countries = (rng % 7).map({0: "United Kingdom", 1: "EIRE", 2: "France", 3: "Germany", 4: "Spain", 5: "Netherlands", 6: "Belgium"})
    return pd.DataFrame(
        {
            "Invoice": (rng % (n // 10 + 1)).astype(str),
            "StockCode": "S" + (rng % 500).astype(str),
            "Description": "DESC-" + (rng % 50).astype(str),
            "Quantity": (rng % 20) + 1,
            "InvoiceDate": pd.Timestamp("2011-01-01") + pd.to_timedelta(rng % 365, unit="D"),
            "Price": ((rng % 100) + 1) * 0.5,
            "Customer ID": (rng % 1000).astype(float),
            "Country": countries,
        }
    )


def _time(fn, repeats: int = 3) -> float:
    best = float("inf")
    for _ in range(repeats):
        t0 = time.perf_counter()
        fn()
        best = min(best, time.perf_counter() - t0)
    return best


def _business_key_naive(df: pd.DataFrame) -> pd.Series:
    return df[BUSINESS_KEY_COLS].astype("string").fillna("<NA>").agg("||".join, axis=1)


def _desc_mode_naive(out: pd.DataFrame) -> pd.Series:
    return (
        out.loc[out["Description"].notna() & (out["Description"].str.len() > 0)]
        .groupby("StockCode")["Description"]
        .agg(lambda s: s.value_counts().index[0] if len(s) else pd.NA)
    )


def bench_cpu(n: int) -> dict:
    df = _synth(n)
    # Active-sales view vs copy (filter only)
    mask = df["Quantity"] > 0

    t_key_naive = _time(lambda: _business_key_naive(df))
    t_key_opt = _time(lambda: _business_key(df))
    t_desc_naive = _time(lambda: _desc_mode_naive(df))
    t_desc_opt = _time(lambda: _description_mode_by_stock(df))
    t_copy = _time(lambda: df.loc[mask].copy())
    t_view = _time(lambda: df.loc[mask])

    def reduction(before: float, after: float) -> float:
        return round(100.0 * (before - after) / before, 1) if before > 0 else 0.0

    return {
        "rows": n,
        "business_key_agg_join_sec": round(t_key_naive, 4),
        "business_key_str_cat_sec": round(t_key_opt, 4),
        "business_key_reduction_pct": reduction(t_key_naive, t_key_opt),
        "description_mode_python_agg_sec": round(t_desc_naive, 4),
        "description_mode_vectorized_sec": round(t_desc_opt, 4),
        "description_mode_reduction_pct": reduction(t_desc_naive, t_desc_opt),
        "active_sales_copy_sec": round(t_copy, 4),
        "active_sales_view_sec": round(t_view, 4),
        "active_sales_reduction_pct": reduction(t_copy, t_view),
    }


def bench_io(tmp: Path, n_files: int = 8, payload_mb: float = 2.0) -> dict:
    """Compare sequential vs ThreadPool writes of independent blobs."""
    tmp.mkdir(parents=True, exist_ok=True)
    blob = b"x" * int(payload_mb * 1_000_000)
    paths = [tmp / f"blob_{i}.bin" for i in range(n_files)]

    def sequential() -> None:
        for p in paths:
            p.write_bytes(blob)

    def threaded() -> None:
        def one(p: Path) -> None:
            p.write_bytes(blob)

        with ThreadPoolExecutor(max_workers=min(4, n_files)) as pool:
            futs = [pool.submit(one, p) for p in paths]
            for f in as_completed(futs):
                f.result()

    # Warm / clear
    for p in paths:
        if p.exists():
            p.unlink()
    t_seq = _time(sequential, repeats=2)
    for p in paths:
        p.unlink(missing_ok=True)
    t_thr = _time(threaded, repeats=2)
    # Also exercise project helper
    for p in paths:
        p.unlink(missing_ok=True)
    t_helper = _time(
        lambda: _run_io_tasks([lambda p=p: p.write_bytes(blob) for p in paths]),
        repeats=2,
    )
    before = t_seq
    after = min(t_thr, t_helper)
    return {
        "files": n_files,
        "payload_mb_each": payload_mb,
        "sequential_write_sec": round(t_seq, 4),
        "threadpool_write_sec": round(t_thr, 4),
        "run_io_tasks_sec": round(t_helper, 4),
        "io_reduction_pct": round(100.0 * (before - after) / before, 1) if before else 0.0,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=100_000)
    parser.add_argument(
        "--out",
        type=Path,
        default=ROOT / "output" / "audit" / "perf_bench_latest.json",
    )
    args = parser.parse_args()

    cpu = bench_cpu(args.rows)
    io = bench_io(ROOT / "output" / "audit" / "_perf_io_tmp")
    payload = {
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "peak_rss_mb": _peak_rss_mb(),
        "cpu": cpu,
        "io": io,
        "notes": (
            "Best-of repeats with time.perf_counter(). "
            "Excel sheet parse stays sequential (one ExcelFile); "
            "ThreadPool is for independent file writes only."
        ),
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(json.dumps(payload, indent=2))
    print(f"\nWrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
