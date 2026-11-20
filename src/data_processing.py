"""Load, audit, and prepare Online Retail II transactions.

Performance notes
-----------------
Prefer vectorized pandas and shared read-only sales views over CPU thread pools.
Excel sheet parse stays sequential (one ``ExcelFile``; openpyxl is not a safe
parallel target). Independent **file writes** (parquet + removed CSV, audit JSON
set) use ``ThreadPoolExecutor`` for overlapping I/O wait.
Excel COM (xlwings) and Outlook (appscript) must stay single-threaded.
"""

from __future__ import annotations

import hashlib
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable, Final

import pandas as pd

from src.config import (
    AUDIT_DIR,
    CLEAN_PARQUET_NAME,
    DATA_PROCESSED,
    DATA_RAW,
    INPUT_XLSX_NAME,
    PREP_SUMMARY_JSON,
    ensure_dirs,
)
from src.constants import Artifacts, BusinessRules, HashingConst, ReturnLineClass
from src.logging_utils import get_logger, timed_phase
from src.metrics import (
    CANCELLATION_INVOICE_PREFIX,
    HOME_COUNTRY,
    METRICS_VERSION,
    NON_PRODUCT_STOCK_CODES,
)

logger = get_logger()


class StaleCleanCacheError(ValueError):
    """Raised when --reuse-clean parquet does not match the current enrich() schema."""

EXCEL_MAX_DATA_ROWS = 1_048_575
BUSINESS_KEY_COLS = [
    "Invoice",
    "StockCode",
    "Description",
    "Quantity",
    "InvoiceDate",
    "Price",
    "Customer ID",
    "Country",
]

WEEKDAY_LABELS = {
    0: "1-Mon",
    1: "2-Tue",
    2: "3-Wed",
    3: "4-Thu",
    4: "5-Fri",
    5: "6-Sat",
    6: "7-Sun",
}


@dataclass
class PrepResult:
    """Prepared clean ledger plus audit metadata and input fingerprint."""

    clean: pd.DataFrame
    audit_summary: dict
    input_path: str
    input_sha256: str


def find_input_workbook(explicit: Path | None = None) -> Path:
    """Resolve the Online Retail II workbook path (explicit, default, or glob under data/raw)."""
    if explicit is not None:
        path = Path(explicit)
        if not path.exists():
            raise FileNotFoundError(path)
        return path
    direct = DATA_RAW / INPUT_XLSX_NAME
    if direct.exists():
        return direct
    matches = list(DATA_RAW.glob("**/online_retail_II.xlsx"))
    if matches:
        return matches[0]
    raise FileNotFoundError(
        f"Missing {INPUT_XLSX_NAME} under {DATA_RAW}. Download Online Retail II from UCI."
    )


def sha256_file(path: Path) -> str:
    """SHA-256 hex digest of a file (chunked read for large workbooks)."""
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(HashingConst.SHA256_CHUNK_BYTES), b""):
            h.update(chunk)
    return h.hexdigest()


def load_raw(path: Path) -> pd.DataFrame:
    """
    Load every worksheet into one frame with provenance columns.

    Adds ``source_sheet``, ``source_row`` (Excel 1-based data row), and ``load_row_id``.

    Uses one ``ExcelFile`` passed into ``read_excel`` (avoids re-opening the
    workbook per sheet). Sheet parsing stays sequential: openpyxl/ExcelFile is
    not a safe parallel target, and threads would still fight the GIL on parse.
    """
    xl = pd.ExcelFile(path)
    sheet_dtype = {
        "Invoice": "string",
        "StockCode": "string",
        "Description": "string",
        "Country": "string",
    }
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, dtype=sheet_dtype)
        # Preserve original column values; normalize names lightly
        rename = {}
        if "Customer ID" not in df.columns and "CustomerID" in df.columns:
            rename["CustomerID"] = "Customer ID"
        if rename:
            df = df.rename(columns=rename)
        missing_cols = [c for c in BUSINESS_KEY_COLS if c not in df.columns]
        if missing_cols:
            raise ValueError(
                f"Sheet {sheet!r} missing required Online Retail II columns: {missing_cols}"
            )
        df["source_sheet"] = sheet
        df["source_row"] = range(2, len(df) + 2)
        frames.append(df)
    out = pd.concat(frames, ignore_index=True)
    out["load_row_id"] = range(1, len(out) + 1)
    logger.info(
        "Loaded raw workbook | rows=%s sheets=%s",
        f"{len(out):,}",
        xl.sheet_names,
    )
    return out


def _business_key(df: pd.DataFrame) -> pd.Series:
    """Stable cross-sheet duplicate key from BUSINESS_KEY_COLS (vectorized str.cat)."""
    keyed = df[BUSINESS_KEY_COLS].astype("string").fillna("<NA>")
    out = keyed.iloc[:, 0]
    for i in range(1, keyed.shape[1]):
        out = out.str.cat(keyed.iloc[:, i], sep="||")
    return out


def apply_cross_sheet_dedupe(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """
    Remove earlier-sheet clones when the same business key appears on multiple sheets.

    Keeps Year 2010-2011 (or lexicographically later sheet) rows.
    Within-sheet exact duplicates are retained and only flagged.

    Also returns a multiset-consistency report: for overlapping keys, compare
    per-sheet multiplicity. On Online Retail II these counts are equal, so
    "keep later sheet" matches multiset max(count_A, count_B).
    """
    work = df.copy()
    work["_biz_key"] = _business_key(work)
    sheet_n = work.groupby("_biz_key")["source_sheet"].nunique()
    overlap_keys = set(sheet_n[sheet_n > 1].index)

    # Multiset consistency for overlapping signatures
    counts = (
        work.loc[work["_biz_key"].isin(overlap_keys)]
        .groupby(["_biz_key", "source_sheet"])
        .size()
        .unstack(fill_value=0)
    )
    unequal = 0
    multiset_kept = 0
    later_kept = 0
    if not counts.empty:
        unequal = int((counts.nunique(axis=1) > 1).sum())
        multiset_kept = int(counts.max(axis=1).sum())
        later_sheet = sorted(work["source_sheet"].unique())[-1]
        if later_sheet in counts.columns:
            later_kept = int(counts[later_sheet].sum())
        else:
            later_kept = multiset_kept

    # Prefer keeping the later worksheet name
    sheet_rank = {s: i for i, s in enumerate(sorted(work["source_sheet"].unique()))}
    work["_sheet_rank"] = work["source_sheet"].map(sheet_rank)
    max_rank = work.groupby("_biz_key")["_sheet_rank"].transform("max")
    drop_mask = work["_biz_key"].isin(overlap_keys) & (work["_sheet_rank"] < max_rank)

    removed_cols = [
        "load_row_id",
        "source_sheet",
        "source_row",
        *BUSINESS_KEY_COLS,
    ]
    removed = work.loc[drop_mask, removed_cols].copy()
    removed["removal_reason"] = "cross_sheet_overlap_earlier_sheet_clone"
    kept = work.loc[~drop_mask].drop(columns=["_sheet_rank"]).copy()

    # Within-sheet exact dup flags (not removed)
    within = kept.groupby(["source_sheet", "_biz_key"]).size()
    within_dup_keys = within[within > 1].index
    kept["exact_dup_within_sheet"] = kept.set_index(["source_sheet", "_biz_key"]).index.isin(
        within_dup_keys
    )
    kept["cross_sheet_overlap_key"] = kept["_biz_key"].isin(overlap_keys)
    kept = kept.drop(columns=["_biz_key"])

    dedupe_audit = {
        "overlap_keys": int(len(overlap_keys)),
        "overlap_keys_unequal_sheet_counts": unequal,
        "overlap_rows_kept_later_sheet": later_kept,
        "overlap_rows_kept_multiset_max": multiset_kept,
        "matches_multiset_max": bool(later_kept == multiset_kept and unequal == 0),
        "policy": (
            "Keep all copies on the lexicographically later source_sheet for overlapping "
            "business keys; retain within-sheet exact duplicates with a flag."
        ),
    }
    return kept, removed, dedupe_audit


def _build_changed_audit(clean: pd.DataFrame) -> pd.DataFrame:
    """Rows whose source-facing values were normalized during enrich (not deleted)."""
    cols = [
        "load_row_id",
        "source_sheet",
        "source_row",
        "field",
        "before",
        "after",
        "change_reason",
    ]
    if "StockCode_raw" not in clean.columns or "StockCode" not in clean.columns:
        return pd.DataFrame(columns=cols)
    changed = clean["StockCode_raw"].fillna("") != clean["StockCode"].fillna("")
    if not changed.any():
        return pd.DataFrame(columns=cols)
    subset = clean.loc[changed, ["load_row_id", "source_sheet", "source_row", "StockCode_raw", "StockCode"]]
    out = subset.rename(columns={"StockCode_raw": "before", "StockCode": "after"})
    out.insert(3, "field", "StockCode")
    out["change_reason"] = "strip_whitespace"
    return out.loc[:, cols].reset_index(drop=True)


def _enrich_types_and_stock(out: pd.DataFrame) -> pd.DataFrame:
    """Cast core types; keep ``StockCode_raw`` and strip ``StockCode`` for analysis."""
    out["Invoice"] = out["Invoice"].astype("string")
    # Keep raw StockCode for traceability; analysis uses stripped StockCode.
    if "StockCode_raw" not in out.columns:
        out["StockCode_raw"] = out["StockCode"].astype("string")
    else:
        out["StockCode_raw"] = out["StockCode_raw"].astype("string")
    out["StockCode"] = out["StockCode_raw"].str.strip()
    out["Description"] = out["Description"].astype("string")
    out["Country"] = out["Country"].astype("string")

    out["InvoiceDate"] = pd.to_datetime(out["InvoiceDate"], errors="coerce")
    out["Quantity"] = pd.to_numeric(out["Quantity"], errors="coerce")
    out["Price"] = pd.to_numeric(out["Price"], errors="coerce")
    return out


def _enrich_classification_flags(out: pd.DataFrame) -> pd.DataFrame:
    """
    Cancellation / return-line taxonomy flags.

    Writes: is_cancellation, is_negative_qty, is_return_line, is_confirmed_return,
    is_stock_adjustment, return_line_class.
    """
    out["is_cancellation"] = (
        out["Invoice"]
        .fillna("")
        .str.upper()
        .str.startswith(CANCELLATION_INVOICE_PREFIX)
    )
    out["is_negative_qty"] = out["Quantity"] < 0
    # Assignment "return indicator": any non-cancellation negative-quantity line.
    # Never use a bare is_return flag — always parent + subtype.
    out["is_return_line"] = out["is_negative_qty"] & (~out["is_cancellation"])
    out["is_confirmed_return"] = out["is_return_line"] & (out["Price"] > 0)
    out["is_stock_adjustment"] = out["is_return_line"] & (~out["is_confirmed_return"])
    out["return_line_class"] = pd.Series(pd.NA, index=out.index, dtype="string")
    out.loc[out["is_confirmed_return"], "return_line_class"] = ReturnLineClass.CONFIRMED.value
    out.loc[out["is_stock_adjustment"], "return_line_class"] = (
        ReturnLineClass.ADJUSTMENT.value
    )
    return out


def _enrich_revenue_columns(out: pd.DataFrame) -> pd.DataFrame:
    """
    Line and aggregate revenue/unit measures used by metrics and pivots.

    Writes: line_revenue, gross_units, returned_units, adjustment_units,
    gross_revenue, return_value, net_revenue_line.
    """
    out["line_revenue"] = out["Quantity"] * out["Price"]
    out["gross_units"] = out["Quantity"].where(
        (out["Quantity"] > 0) & (~out["is_cancellation"]), 0
    ).fillna(0)
    out["returned_units"] = (
        (-out["Quantity"]).where(out["is_confirmed_return"], 0).fillna(0)
    )
    out["adjustment_units"] = (
        (-out["Quantity"]).where(out["is_stock_adjustment"], 0).fillna(0)
    )
    out["gross_revenue"] = out["line_revenue"].where(
        (out["Quantity"] > 0) & (~out["is_cancellation"]), 0.0
    ).fillna(0.0)
    out["return_value"] = (
        (-out["line_revenue"]).where(out["is_confirmed_return"], 0.0).fillna(0.0)
    )
    out["net_revenue_line"] = out["gross_revenue"] - out["return_value"]
    return out


def _enrich_time_columns(out: pd.DataFrame) -> pd.DataFrame:
    """Writes year_month, hour, weekday_label from InvoiceDate."""
    # Period→str is faster than strftime on ~1M rows and yields the same YYYY-MM.
    periods = out["InvoiceDate"].dt.to_period("M")
    out["year_month"] = periods.astype(str).where(out["InvoiceDate"].notna(), other=pd.NA)
    out["hour"] = out["InvoiceDate"].dt.hour
    out["weekday_label"] = out["InvoiceDate"].dt.day_of_week.map(WEEKDAY_LABELS)
    return out


def _description_mode_by_stock(out: pd.DataFrame) -> pd.Series:
    """Most frequent non-empty Description per StockCode (ties → earliest row)."""
    mask = out["Description"].notna() & (out["Description"].str.len() > 0)
    if not mask.any():
        return pd.Series(dtype="string")
    valid = out.loc[mask, ["StockCode", "Description"]].copy()
    valid["_i"] = range(len(valid))
    valid["_n"] = valid.groupby(["StockCode", "Description"], sort=False)[
        "Description"
    ].transform("size")
    ranked = valid.sort_values(["StockCode", "_n", "_i"], ascending=[True, False, True])
    return ranked.drop_duplicates("StockCode").set_index("StockCode")["Description"]


def _enrich_non_product_and_labels(out: pd.DataFrame) -> pd.DataFrame:
    """
    Non-product flag, canonical description/label, Country_NonUK, cancellation measures.

    Writes: is_non_product, Description_std, product_label, Country_NonUK,
    cancellation_value_line, cancelled_units.

    ``is_non_product`` is True only for codes in ``NON_PRODUCT_STOCK_CODES``
    (case-insensitive). Letter-only catalogue SKUs are kept as products.
    """
    stock_upper = out["StockCode"].str.upper()
    out["is_non_product"] = stock_upper.isin(
        {s.upper() for s in NON_PRODUCT_STOCK_CODES}
    )

    desc_mode = _description_mode_by_stock(out)
    out["Description_std"] = out["StockCode"].map(desc_mode).fillna(out["Description"])
    out["product_label"] = (
        out["StockCode"].astype("string").fillna("")
        + " — "
        + out["Description_std"].astype("string").fillna("")
    )
    # Country pivot row field: home market blanked so it is not listed as a row.
    out["Country_NonUK"] = out["Country"].where(
        out["Country"].fillna("") != HOME_COUNTRY, other=pd.NA
    )
    out["cancellation_value_line"] = (
        (-out["line_revenue"]).where(out["is_cancellation"], 0.0).clip(lower=0).fillna(0.0)
    )
    out["cancelled_units"] = (
        (-out["Quantity"]).where(out["is_cancellation"], 0).clip(lower=0).fillna(0)
    )
    return out


def _mark_first_in_group(
    out: pd.DataFrame, mask: pd.Series, subset: list[str], col: str
) -> None:
    """Set ``col`` to 1 on the first row of each ``subset`` group within ``mask``."""
    out[col] = 0
    if not mask.any():
        return
    sub_idx = out.index[mask]
    first = ~out.loc[sub_idx].duplicated(subset=subset, keep="first")
    out.loc[sub_idx[first.to_numpy()], col] = 1


def _enrich_invoice_helper_flags(out: pd.DataFrame) -> pd.DataFrame:
    """
    One-hot-style helpers so PivotTables can Sum distinct invoices/customers.

    Writes inv_flag_by_{product,country,month,customer} and
    cust_flag_by_{country,month} (1 on first occurrence within group).
    """
    active = ~out["is_cancellation"].fillna(False)
    _mark_first_in_group(out, active, ["StockCode", "Invoice"], "inv_flag_by_product")
    _mark_first_in_group(out, active, ["Country", "Invoice"], "inv_flag_by_country")
    _mark_first_in_group(out, active, ["year_month", "Invoice"], "inv_flag_by_month")

    known_cust = active & out["Customer ID"].notna()
    _mark_first_in_group(
        out, known_cust, ["Customer ID", "Invoice"], "inv_flag_by_customer"
    )
    _mark_first_in_group(
        out, known_cust, ["Country", "Customer ID"], "cust_flag_by_country"
    )
    _mark_first_in_group(
        out, known_cust, ["year_month", "Customer ID"], "cust_flag_by_month"
    )
    return out


def _enrich_scope_helpers(out: pd.DataFrame) -> pd.DataFrame:
    """
    Scoped revenue/flag columns so pivots can filter UK / fees / anonymous IDs
    without copying Clean_Data onto each pivot sheet.
    """
    non_uk = out["Country"].fillna("") != HOME_COUNTRY
    out["gross_revenue_non_uk"] = out["gross_revenue"].where(non_uk, 0.0)
    out["inv_flag_by_country_non_uk"] = out["inv_flag_by_country"].where(non_uk, 0)
    out["cust_flag_by_country_non_uk"] = out["cust_flag_by_country"].where(non_uk, 0)

    prod = ~out["is_non_product"].fillna(False)
    out["gross_revenue_product"] = out["gross_revenue"].where(prod, 0.0)
    out["gross_units_product"] = out["gross_units"].where(prod, 0)
    out["returned_units_product"] = out["returned_units"].where(prod, 0)
    out["adjustment_units_product"] = out["adjustment_units"].where(prod, 0)
    out["return_value_product"] = out["return_value"].where(prod, 0.0)
    out["inv_flag_by_product_only"] = out["inv_flag_by_product"].where(prod, 0)

    known = out["Customer ID"].notna()
    out["gross_revenue_known_customer"] = out["gross_revenue"].where(known, 0.0)
    return out


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add classification, revenue, time, and pivot-helper columns to a raw/kept ledger.

    Does not drop rows. Column families are applied via ``_enrich_*`` helpers in order.
    Output must include every name in ``ENRICH_REQUIRED_COLUMNS``.
    """
    out = df.copy()
    _enrich_types_and_stock(out)
    _enrich_classification_flags(out)
    _enrich_revenue_columns(out)
    _enrich_time_columns(out)
    _enrich_non_product_and_labels(out)
    _enrich_invoice_helper_flags(out)
    _enrich_scope_helpers(out)
    return out


# Columns that a clean parquet must have after the current enrich() (+ load/dedupe
# provenance). Used by --reuse-clean so schema drift fails loudly instead of
# silently analysing a stale cache. Keep in sync with enrich() / _enrich_* helpers.
ENRICH_REQUIRED_COLUMNS: Final[tuple[str, ...]] = (
    "load_row_id",
    "source_sheet",
    "source_row",
    "Invoice",
    "StockCode_raw",
    "StockCode",
    "Description",
    "Description_std",
    "product_label",
    "Quantity",
    "InvoiceDate",
    "Price",
    "Customer ID",
    "Country",
    "Country_NonUK",
    "line_revenue",
    "gross_revenue",
    "return_value",
    "net_revenue_line",
    "gross_units",
    "returned_units",
    "adjustment_units",
    "cancellation_value_line",
    "cancelled_units",
    "year_month",
    "weekday_label",
    "hour",
    "is_cancellation",
    "is_negative_qty",
    "is_return_line",
    "is_confirmed_return",
    "is_stock_adjustment",
    "return_line_class",
    "is_non_product",
    "exact_dup_within_sheet",
    "cross_sheet_overlap_key",
    "inv_flag_by_product",
    "inv_flag_by_country",
    "inv_flag_by_month",
    "inv_flag_by_customer",
    "cust_flag_by_country",
    "cust_flag_by_month",
    "gross_revenue_non_uk",
    "inv_flag_by_country_non_uk",
    "cust_flag_by_country_non_uk",
    "gross_revenue_product",
    "gross_units_product",
    "returned_units_product",
    "adjustment_units_product",
    "return_value_product",
    "inv_flag_by_product_only",
    "gross_revenue_known_customer",
)


def validate_enrich_schema(df: pd.DataFrame) -> list[str]:
    """
    Return error strings if ``df`` is missing columns required by the current enrich().

    Empty list means the frame is schema-compatible with downstream analyses.
    """
    missing = [c for c in ENRICH_REQUIRED_COLUMNS if c not in df.columns]
    if not missing:
        return []
    return [
        "Cached clean parquet is missing columns expected by the current enrich() "
        f"schema: {missing}. Re-run without --reuse-clean."
    ]


def _missing_value_profile(clean: pd.DataFrame) -> dict:
    """Null counts/shares for core business columns present on clean."""
    return {
        col: {
            "null_rows": int(clean[col].isna().sum()),
            "null_share": float(clean[col].isna().mean()),
        }
        for col in (
            "Invoice",
            "StockCode",
            "Description",
            "Quantity",
            "InvoiceDate",
            "Price",
            "Customer ID",
            "Country",
        )
        if col in clean.columns
    }


def _classification_counts(clean: pd.DataFrame, missing_cust: pd.Series) -> dict:
    """Flag / edge-case row counts for the data-quality audit."""
    return {
        "cancellations": int(clean["is_cancellation"].sum()),
        "return_lines": int(clean["is_return_line"].sum()),
        "confirmed_returns": int(clean["is_confirmed_return"].sum()),
        "stock_adjustments": int(clean["is_stock_adjustment"].sum()),
        "negative_qty": int(clean["is_negative_qty"].sum()),
        "within_sheet_exact_dup_rows": int(clean["exact_dup_within_sheet"].sum()),
        "zero_price_rows": int((clean["Price"] == 0).sum()),
        "negative_price_rows": int((clean["Price"] < 0).sum()),
        "zero_quantity_rows": int((clean["Quantity"] == 0).sum()),
        "missing_customer_id_rows": int(missing_cust.sum()),
        "cancellation_with_positive_qty": int(
            (clean["is_cancellation"] & (clean["Quantity"] > 0)).sum()
        ),
    }


def build_data_quality_summary(clean: pd.DataFrame, audit: dict) -> dict:
    """
    Build the assignment data-quality audit JSON.

    Includes row reconciliation, return-line taxonomy identity, missing-value
    profile, and top countries by missing Customer ID rows.
    """
    missing_cust = clean["Customer ID"].isna()
    missing_by_country = (
        clean.loc[missing_cust]
        .groupby(clean.loc[missing_cust, "Country"].fillna("Unknown"))
        .size()
        .sort_values(ascending=False)
        .head(10)
        .astype(int)
        .to_dict()
    )
    return {
        "metrics_contract_version": METRICS_VERSION,
        "row_reconciliation": {
            "raw_rows": audit.get("raw_rows"),
            "removed_cross_sheet_clones": audit.get("removed_cross_sheet_clones"),
            "clean_rows": audit.get("clean_rows"),
            "excel_max_data_rows": EXCEL_MAX_DATA_ROWS,
            "excel_rows_remaining_incl_header": int(
                EXCEL_MAX_DATA_ROWS - int(audit.get("clean_rows", 0))
            ),
            "identity_check_raw_eq_clean_plus_removed": bool(
                audit.get("raw_rows")
                == int(audit.get("clean_rows", 0))
                + int(audit.get("removed_cross_sheet_clones", 0))
            ),
        },
        "dedupe": audit.get("dedupe", {}),
        "classification_counts": _classification_counts(clean, missing_cust),
        "return_line_taxonomy": {
            "assignment_return_indicator": "is_return_line",
            "meaning": (
                "is_return_line = Quantity<0 and not cancellation. "
                "Every return line is exactly one of: is_confirmed_return (Price>0) "
                "or is_stock_adjustment (else). Do not read a bare is_return flag."
            ),
            "identity_return_lines_eq_confirmed_plus_adjustments": bool(
                int(clean["is_return_line"].sum())
                == int(clean["is_confirmed_return"].sum())
                + int(clean["is_stock_adjustment"].sum())
            ),
        },
        "missing_values": _missing_value_profile(clean),
        "missing_customer_id_rows_by_country_top10": missing_by_country,
    }


def build_revenue_reconciliation(clean: pd.DataFrame) -> dict:
    """
    Prove Gross / confirmed-return / cancellation / Net identities on the clean ledger.

    Stock adjustments contribute units only (not return_value / Net).
    """
    sales = clean.loc[~clean["is_cancellation"]]
    canc = clean.loc[clean["is_cancellation"]]
    gross = float(sales["gross_revenue"].sum())
    return_value = float(sales["return_value"].sum())
    cancellation_value = float((-canc["line_revenue"]).clip(lower=0).sum())
    adjustment_units = float(clean["adjustment_units"].sum())
    confirmed_returned_units = float(clean["returned_units"].sum())
    net = gross - return_value
    net_after_c = net - cancellation_value
    return {
        "gross_sales": gross,
        "confirmed_return_value": return_value,
        "confirmed_returned_units": confirmed_returned_units,
        "stock_adjustment_units": adjustment_units,
        "cancellation_value": cancellation_value,
        "net_sales": net,
        "net_after_cancellations": net_after_c,
        "identities": {
            "net_eq_gross_minus_confirmed_returns": abs(net - (gross - return_value))
            < BusinessRules.FLOAT_EQ_TOLERANCE,
            "net_after_c_eq_net_minus_cancellation": abs(
                net_after_c - (net - cancellation_value)
            )
            < BusinessRules.FLOAT_EQ_TOLERANCE,
        },
        "notes": (
            "Stock adjustments (non-C, Price<=0 negatives) do not reduce net_sales cash "
            "because return_value only includes Price>0 confirmed returns."
        ),
    }


def _build_prep_audit_summary(
    *,
    path: Path,
    digest: str,
    raw_rows: int,
    removed: pd.DataFrame,
    clean: pd.DataFrame,
    dedupe_audit: dict,
) -> dict:
    """Assemble the prep_summary.json payload (counts, paths, return-line note)."""
    dates = clean["InvoiceDate"]
    return {
        "input_path": str(path),
        "input_sha256": digest,
        "raw_rows": raw_rows,
        "removed_cross_sheet_clones": int(len(removed)),
        "clean_rows": int(len(clean)),
        "excel_max_data_rows": EXCEL_MAX_DATA_ROWS,
        "fits_excel": bool(len(clean) <= EXCEL_MAX_DATA_ROWS),
        "within_sheet_exact_dup_rows": int(clean["exact_dup_within_sheet"].sum()),
        "cancellations": int(clean["is_cancellation"].sum()),
        "return_lines": int(clean["is_return_line"].sum()),
        "confirmed_returns": int(clean["is_confirmed_return"].sum()),
        "stock_adjustments": int(clean["is_stock_adjustment"].sum()),
        "missing_customer_id": int(clean["Customer ID"].isna().sum()),
        "date_min": str(dates.min()),
        "date_max": str(dates.max()),
        "dedupe": dedupe_audit,
        "enrich_schema_columns": list(ENRICH_REQUIRED_COLUMNS),
        "removed_audit_file": f"audit/{Artifacts.REMOVED_CLONES_CSV}",
        "changed_audit_file": f"audit/{Artifacts.CHANGED_RECORDS_CSV}",
        "clean_parquet": f"data/processed/{CLEAN_PARQUET_NAME}",
        "data_quality_file": f"audit/{Artifacts.DATA_QUALITY_JSON}",
        "revenue_reconciliation_file": f"audit/{Artifacts.REVENUE_RECON_JSON}",
        "return_line_note": (
            "Assignment return indicator = is_return_line "
            f"({int(clean['is_return_line'].sum()):,} rows) = "
            f"is_confirmed_return ({int(clean['is_confirmed_return'].sum()):,}) + "
            f"is_stock_adjustment ({int(clean['is_stock_adjustment'].sum()):,})."
        ),
    }


def _run_io_tasks(tasks: list[Callable[[], None]], *, max_workers: int = 4) -> None:
    """Run independent I/O callables concurrently; surface the first exception."""
    if not tasks:
        return
    workers = max(1, min(max_workers, len(tasks)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(task) for task in tasks]
        for fut in as_completed(futures):
            fut.result()


def _write_prep_audit_artifacts(clean: pd.DataFrame, audit_summary: dict) -> int:
    """Write changed-records CSV + quality/revenue/prep JSON audits. Returns changed count."""
    quality = build_data_quality_summary(clean, audit_summary)
    revenue = build_revenue_reconciliation(clean)
    changed = _build_changed_audit(clean)
    audit_summary["changed_records"] = int(len(changed))
    audit_summary["revenue_reconciliation"] = revenue

    changed_path = AUDIT_DIR / Artifacts.CHANGED_RECORDS_CSV
    prep_text = json.dumps(audit_summary, indent=2)
    quality_text = json.dumps(quality, indent=2)
    revenue_text = json.dumps(revenue, indent=2)

    # Four independent file writes — overlap disk wait via a small thread pool.
    _run_io_tasks(
        [
            lambda: changed.to_csv(changed_path, index=False),
            lambda: (AUDIT_DIR / PREP_SUMMARY_JSON).write_text(
                prep_text, encoding="utf-8"
            ),
            lambda: (AUDIT_DIR / Artifacts.DATA_QUALITY_JSON).write_text(
                quality_text, encoding="utf-8"
            ),
            lambda: (AUDIT_DIR / Artifacts.REVENUE_RECON_JSON).write_text(
                revenue_text, encoding="utf-8"
            ),
        ]
    )
    return int(len(changed))


def prepare_dataset(input_path: Path | None = None) -> PrepResult:
    """
    Full prep pipeline: locate input → load → dedupe → enrich → persist parquet/audits.

    Raises if clean rows would exceed Excel's worksheet limit (no silent subset).
    """
    ensure_dirs()
    path = find_input_workbook(input_path)

    with timed_phase("hash_input", path=str(path)) as meta:
        digest = sha256_file(path)
        meta["sha256_prefix"] = digest[:12]

    with timed_phase("load_raw", path=str(path)) as meta:
        raw = load_raw(path)
        raw_rows = len(raw)
        meta["rows"] = raw_rows

    with timed_phase("cross_sheet_dedupe", rows=raw_rows) as meta:
        kept, removed, dedupe_audit = apply_cross_sheet_dedupe(raw)
        meta["kept"] = len(kept)
        meta["removed"] = len(removed)

    with timed_phase("enrich", rows=len(kept)) as meta:
        clean = enrich(kept)
        meta["clean_rows"] = len(clean)

    if len(clean) > EXCEL_MAX_DATA_ROWS:
        raise RuntimeError(
            f"Clean_Data would have {len(clean):,} rows, exceeding Excel limit "
            f"{EXCEL_MAX_DATA_ROWS:,}. Refusing silent subset fallback."
        )

    with timed_phase("persist_clean", rows=len(clean)) as meta:
        parquet = DATA_PROCESSED / CLEAN_PARQUET_NAME
        removed_path = AUDIT_DIR / Artifacts.REMOVED_CLONES_CSV
        # Parquet + removed-clones CSV are independent sinks — overlap I/O.
        _run_io_tasks(
            [
                lambda: clean.to_parquet(parquet, index=False),
                lambda: removed.to_csv(removed_path, index=False),
            ],
            max_workers=2,
        )
        meta["parquet"] = str(parquet)
        meta["removed_csv"] = str(removed_path)

    schema_errors = validate_enrich_schema(clean)
    if schema_errors:
        raise RuntimeError(schema_errors[0])

    audit_summary = _build_prep_audit_summary(
        path=path,
        digest=digest,
        raw_rows=raw_rows,
        removed=removed,
        clean=clean,
        dedupe_audit=dedupe_audit,
    )

    with timed_phase("write_prep_audit", clean_rows=len(clean)) as meta:
        changed_n = _write_prep_audit_artifacts(clean, audit_summary)
        meta["changed_records"] = changed_n

    if not dedupe_audit.get("matches_multiset_max", True):
        logger.warning(
            "Cross-sheet dedupe unequal | policy=later_sheet_keep "
            "note=does_not_match_multiset_max"
        )
    logger.info(
        "Prepared clean data | rows=%s removed_clones=%s changed=%s "
        "return_lines=%s confirmed_returns=%s stock_adjustments=%s",
        f"{len(clean):,}",
        f"{len(removed):,}",
        f"{changed_n:,}",
        f"{int(clean['is_return_line'].sum()):,}",
        f"{int(clean['is_confirmed_return'].sum()):,}",
        f"{int(clean['is_stock_adjustment'].sum()):,}",
    )
    return PrepResult(
        clean=clean,
        audit_summary=audit_summary,
        input_path=str(path),
        input_sha256=digest,
    )


# Excel Clean_Data sheet columns (subset of ENRICH_REQUIRED_COLUMNS; lean for pivots).
CLEAN_DATA_COLUMNS = [
    "load_row_id",
    "source_sheet",
    "source_row",
    "Invoice",
    "StockCode_raw",
    "StockCode",
    "Description",
    "Description_std",
    "product_label",
    "Quantity",
    "InvoiceDate",
    "Price",
    "Customer ID",
    "Country",
    "Country_NonUK",
    "line_revenue",
    "gross_revenue",
    "return_value",
    "gross_units",
    "returned_units",
    "adjustment_units",
    "cancellation_value_line",
    "cancelled_units",
    "year_month",
    "weekday_label",
    "hour",
    "is_cancellation",
    "is_negative_qty",
    "is_return_line",
    "is_confirmed_return",
    "is_stock_adjustment",
    "return_line_class",
    "is_non_product",
    "exact_dup_within_sheet",
    "cross_sheet_overlap_key",
    "inv_flag_by_month",
    "inv_flag_by_customer",
    "cust_flag_by_month",
    "gross_revenue_non_uk",
    "inv_flag_by_country_non_uk",
    "cust_flag_by_country_non_uk",
    "gross_revenue_product",
    "gross_units_product",
    "returned_units_product",
    "adjustment_units_product",
    "return_value_product",
    "inv_flag_by_product_only",
    "gross_revenue_known_customer",
]


def load_prepared(parquet_path: Path | None = None) -> PrepResult:
    """
    Reload previously prepared clean parquet + audit summary.

    Validates that the parquet still matches ``ENRICH_REQUIRED_COLUMNS`` so
    ``--reuse-clean`` cannot silently analyse a schema-stale cache.
    """
    ensure_dirs()
    parquet = Path(parquet_path) if parquet_path else DATA_PROCESSED / CLEAN_PARQUET_NAME
    summary_path = AUDIT_DIR / PREP_SUMMARY_JSON
    if not parquet.exists():
        raise FileNotFoundError(parquet)
    if not summary_path.exists():
        raise FileNotFoundError(summary_path)
    clean = pd.read_parquet(parquet)
    schema_errors = validate_enrich_schema(clean)
    if schema_errors:
        raise StaleCleanCacheError(schema_errors[0])
    audit_summary = json.loads(summary_path.read_text(encoding="utf-8"))
    return PrepResult(
        clean=clean,
        audit_summary=audit_summary,
        input_path=str(audit_summary.get("input_path", "")),
        input_sha256=str(audit_summary.get("input_sha256", "")),
    )
