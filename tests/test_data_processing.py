"""Tests for load/dedupe/enrich rules in src/data_processing.py."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.data_processing import (
    CLEAN_DATA_COLUMNS,
    ENRICH_REQUIRED_COLUMNS,
    StaleCleanCacheError,
    apply_cross_sheet_dedupe,
    enrich,
    find_input_workbook,
    load_prepared,
    load_raw,
    prepare_dataset,
    sha256_file,
    validate_enrich_schema,
)


def test_enrich_schema_contract(raw_transactions: pd.DataFrame):
    """enrich() must emit every ENRICH_REQUIRED_COLUMNS name; Excel cols ⊆ enrich."""
    kept, _, _ = apply_cross_sheet_dedupe(raw_transactions)
    out = enrich(kept)
    assert validate_enrich_schema(out) == []
    assert set(ENRICH_REQUIRED_COLUMNS) <= set(out.columns)
    assert set(CLEAN_DATA_COLUMNS) <= set(ENRICH_REQUIRED_COLUMNS)


def test_is_non_product_uses_allowlist_only():
    """Fees/postage match NON_PRODUCT_STOCK_CODES; letter-only catalogue SKUs do not."""
    from src.metrics import NON_PRODUCT_STOCK_CODES

    rows = []
    for i, code in enumerate(
        ["POST", "C2", "B", "DCGSSGIRL", "ABCXYZ", "85123A"], start=1
    ):
        rows.append(
            {
                "load_row_id": i,
                "source_sheet": "Year 2010-2011",
                "source_row": i,
                "Invoice": str(1000 + i),
                "StockCode": code,
                "Description": f"desc-{code}",
                "Quantity": 1,
                "InvoiceDate": "2011-01-01",
                "Price": 1.0,
                "Customer ID": 1.0,
                "Country": "United Kingdom",
                "exact_dup_within_sheet": False,
                "cross_sheet_overlap_key": f"k{i}",
            }
        )
    out = enrich(pd.DataFrame(rows))
    by_code = out.set_index("StockCode")["is_non_product"]
    assert bool(by_code["POST"]) is True
    assert bool(by_code["C2"]) is True
    assert bool(by_code["B"]) is True
    assert bool(by_code["DCGSSGIRL"]) is False
    assert bool(by_code["ABCXYZ"]) is False
    assert bool(by_code["85123A"]) is False
    # Allowlist membership is the sole rule (case-insensitive).
    assert "POST" in {s.upper() for s in NON_PRODUCT_STOCK_CODES}


def test_validate_enrich_schema_detects_missing_column():
    errors = validate_enrich_schema(pd.DataFrame({"Invoice": [1]}))
    assert errors
    assert "missing columns" in errors[0] or "enrich()" in errors[0]


def test_cross_sheet_dedupe_keeps_later_sheet(raw_transactions: pd.DataFrame):
    kept, removed, dedupe = apply_cross_sheet_dedupe(raw_transactions)
    assert len(removed) == 1
    assert removed.iloc[0]["source_sheet"] == "Year 2009-2010"
    assert removed.iloc[0]["removal_reason"] == "cross_sheet_overlap_earlier_sheet_clone"
    assert dedupe["matches_multiset_max"] is True
    # Later sheet clone remains
    assert (
        (kept["source_sheet"] == "Year 2010-2011") & (kept["Invoice"] == "10001")
    ).any()


def test_within_sheet_exact_dups_are_flagged_not_removed(raw_transactions: pd.DataFrame):
    kept, _, _ = apply_cross_sheet_dedupe(raw_transactions)
    dups = kept.loc[
        (kept["Invoice"] == "10002") & (kept["StockCode"] == "P1")
    ]
    assert len(dups) == 2
    assert bool(dups["exact_dup_within_sheet"].all())


def test_enrich_classifies_zero_price_negatives_as_adjustments(raw_transactions: pd.DataFrame):
    kept, _, _ = apply_cross_sheet_dedupe(raw_transactions)
    out = enrich(kept)

    assert "line_revenue" in out.columns
    assert "product_label" in out.columns
    assert "Country_NonUK" in out.columns

    canc = out.loc[out["Invoice"].astype(str).str.startswith("C")]
    assert canc["is_cancellation"].all()
    assert (~canc["is_return_line"]).all()
    assert (~canc["is_confirmed_return"]).all()

    # Price=0 non-C negative → return line + stock adjustment subtype
    adj = out.loc[(out["Quantity"] < 0) & (~out["is_cancellation"]) & (out["Price"] == 0)]
    assert adj["is_return_line"].all()
    assert adj["is_stock_adjustment"].all()
    assert (~adj["is_confirmed_return"]).all()
    assert (adj["return_line_class"] == "stock_adjustment").all()
    assert float(out["return_value"].sum()) == 0.0
    assert int(out["is_return_line"].sum()) == int(out["is_confirmed_return"].sum()) + int(
        out["is_stock_adjustment"].sum()
    )

    # UK blanked in Country_NonUK; EIRE kept
    uk = out.loc[out["Country"] == "United Kingdom", "Country_NonUK"]
    assert pd.isna(uk).all()
    eire = out.loc[out["Country"] == "EIRE", "Country_NonUK"]
    assert (eire.astype(str) == "EIRE").all()

    # Gross revenue excludes cancellations and negative qty
    sales = out.loc[~out["is_cancellation"] & (out["Quantity"] > 0)]
    expected = float((sales["Quantity"] * sales["Price"]).sum())
    assert abs(float(out["gross_revenue"].sum()) - expected) < 1e-9


def test_confirmed_return_requires_positive_price():
    from src.data_processing import enrich

    df = pd.DataFrame(
        [
            {
                "load_row_id": 1,
                "source_sheet": "Year 2010-2011",
                "source_row": 1,
                "Invoice": "R1",
                "StockCode": "X",
                "Description": "ITEM",
                "Quantity": -3,
                "InvoiceDate": "2011-04-01",
                "Price": 2.5,
                "Customer ID": 1.0,
                "Country": "EIRE",
            }
        ]
    )
    out = enrich(df)
    assert bool(out.iloc[0]["is_return_line"])
    assert bool(out.iloc[0]["is_confirmed_return"])
    assert not bool(out.iloc[0]["is_stock_adjustment"])
    assert out.iloc[0]["return_line_class"] == "confirmed_return"
    assert float(out.iloc[0]["returned_units"]) == 3.0
    assert float(out.iloc[0]["return_value"]) == 7.5
    assert "StockCode_raw" in out.columns


def test_stockcode_strip_preserves_raw():
    df = pd.DataFrame(
        [
            {
                "load_row_id": 1,
                "source_sheet": "Year 2010-2011",
                "source_row": 1,
                "Invoice": "1",
                "StockCode": "47503J ",
                "Description": "X",
                "Quantity": 1,
                "InvoiceDate": "2011-01-01",
                "Price": 1.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
            }
        ]
    )
    out = enrich(df)
    assert out.iloc[0]["StockCode_raw"] == "47503J "
    assert out.iloc[0]["StockCode"] == "47503J"


def test_changed_and_removed_audit_shape(raw_transactions: pd.DataFrame):
    from src.data_processing import _build_changed_audit

    kept, removed, _ = apply_cross_sheet_dedupe(raw_transactions)
    assert "Quantity" in removed.columns
    assert "Price" in removed.columns
    clean = enrich(kept)
    changed = _build_changed_audit(clean)
    assert list(changed.columns) == [
        "load_row_id",
        "source_sheet",
        "source_row",
        "field",
        "before",
        "after",
        "change_reason",
    ]


def test_find_input_workbook_explicit_missing_glob(tmp_path, monkeypatch):
    from src import data_processing as dp
    from src import config as cfg

    monkeypatch.setattr(dp, "DATA_RAW", tmp_path)
    monkeypatch.setattr(cfg, "DATA_RAW", tmp_path)

    explicit = tmp_path / "given.xlsx"
    explicit.write_bytes(b"x")
    assert find_input_workbook(explicit) == explicit

    with pytest.raises(FileNotFoundError):
        find_input_workbook(tmp_path / "missing.xlsx")

    with pytest.raises(FileNotFoundError, match="Missing"):
        find_input_workbook(None)

    nested = tmp_path / "sub"
    nested.mkdir()
    target = nested / "online_retail_II.xlsx"
    target.write_bytes(b"y")
    assert find_input_workbook(None) == target

    direct = tmp_path / "online_retail_II.xlsx"
    direct.write_bytes(b"z")
    assert find_input_workbook(None) == direct


def test_sha256_file(tmp_path):
    p = tmp_path / "f.bin"
    p.write_bytes(b"hello")
    digest = sha256_file(p)
    assert len(digest) == 64
    assert digest == sha256_file(p)


def test_load_raw_multi_sheet(tmp_path):
    path = tmp_path / "tiny.xlsx"
    sheet_a = pd.DataFrame(
        [
            {
                "Invoice": "1",
                "StockCode": "A",
                "Description": "x",
                "Quantity": 1,
                "InvoiceDate": "2011-01-01",
                "Price": 1.0,
                "CustomerID": 10.0,
                "Country": "EIRE",
            }
        ]
    )
    sheet_b = pd.DataFrame(
        [
            {
                "Invoice": "2",
                "StockCode": "B",
                "Description": "y",
                "Quantity": 2,
                "InvoiceDate": "2011-02-01",
                "Price": 2.0,
                "Customer ID": 20.0,
                "Country": "France",
            }
        ]
    )
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        sheet_a.to_excel(writer, sheet_name="Year 2009-2010", index=False)
        sheet_b.to_excel(writer, sheet_name="Year 2010-2011", index=False)
    raw = load_raw(path)
    assert len(raw) == 2
    assert "Customer ID" in raw.columns
    assert set(raw["source_sheet"]) == {"Year 2009-2010", "Year 2010-2011"}
    assert list(raw["load_row_id"]) == [1, 2]


def test_load_raw_rejects_missing_business_columns(tmp_path):
    path = tmp_path / "bad.xlsx"
    pd.DataFrame({"Invoice": ["1"], "Quantity": [1]}).to_excel(
        path, sheet_name="Year 2010-2011", index=False
    )
    with pytest.raises(ValueError, match="missing required Online Retail II columns"):
        load_raw(path)


def test_enrich_stockcode_raw_already_present():
    df = pd.DataFrame(
        [
            {
                "load_row_id": 1,
                "source_sheet": "Year 2010-2011",
                "source_row": 1,
                "Invoice": "1",
                "StockCode": "XX ",
                "StockCode_raw": "RAW ",
                "Description": "x",
                "Quantity": 1,
                "InvoiceDate": "2011-01-01",
                "Price": 1.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
            }
        ]
    )
    out = enrich(df)
    assert out.iloc[0]["StockCode_raw"] == "RAW "
    assert out.iloc[0]["StockCode"] == "RAW"


def test_unequal_sheet_counts_dedupe_warning():
    """Overlap key with different multiplicity across sheets → warning audit flag."""
    rows = []
    # Same business key twice on earlier sheet, once on later
    for i, sheet in enumerate(["Year 2009-2010", "Year 2009-2010", "Year 2010-2011"]):
        rows.append(
            {
                "load_row_id": i + 1,
                "source_sheet": sheet,
                "source_row": i + 2,
                "Invoice": "SAME",
                "StockCode": "P",
                "Description": "X",
                "Quantity": 1,
                "InvoiceDate": "2011-01-01",
                "Price": 1.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
            }
        )
    kept, removed, dedupe = apply_cross_sheet_dedupe(pd.DataFrame(rows))
    assert dedupe["overlap_keys_unequal_sheet_counts"] >= 1
    assert dedupe["matches_multiset_max"] is False
    assert len(removed) == 2  # both earlier-sheet copies removed


def test_prepare_dataset_and_load_prepared(tmp_path, monkeypatch, raw_transactions):
    from src import data_processing as dp
    from src import config as cfg
    from src import logging_utils as lu

    raw_dir = tmp_path / "raw"
    processed = tmp_path / "processed"
    audit = tmp_path / "audit"
    logs = tmp_path / "logs"
    for p in (raw_dir, processed, audit, logs):
        p.mkdir()

    monkeypatch.setattr(dp, "DATA_RAW", raw_dir)
    monkeypatch.setattr(dp, "DATA_PROCESSED", processed)
    monkeypatch.setattr(dp, "AUDIT_DIR", audit)
    monkeypatch.setattr(cfg, "DATA_RAW", raw_dir)
    monkeypatch.setattr(cfg, "DATA_PROCESSED", processed)
    monkeypatch.setattr(cfg, "AUDIT_DIR", audit)
    monkeypatch.setattr(cfg, "LOG_DIR", logs)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", tmp_path / "out")
    monkeypatch.setattr(cfg, "PNG_DIR", tmp_path / "out" / "png")
    monkeypatch.setattr(lu, "LOG_DIR", logs)
    (tmp_path / "out").mkdir()
    (tmp_path / "out" / "png").mkdir()

    # Write tiny multi-sheet xlsx from fixture rows (drop helper cols)
    path = raw_dir / "online_retail_II.xlsx"
    cols = [
        "Invoice",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "Price",
        "Customer ID",
        "Country",
    ]
    # Include a StockCode with trailing space so changed_records is non-empty
    frame = raw_transactions.copy()
    frame.loc[frame.index[0], "StockCode"] = "P1 "
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for sheet, g in frame.groupby("source_sheet"):
            g[cols].to_excel(writer, sheet_name=sheet, index=False)

    result = prepare_dataset(path)
    assert len(result.clean) > 0
    assert (processed / "clean_transactions.parquet").exists()
    assert (audit / "prep_summary.json").exists()
    assert (audit / "changed_records.csv").exists()

    reloaded = load_prepared()
    assert len(reloaded.clean) == len(result.clean)
    assert reloaded.input_sha256 == result.input_sha256
    assert result.audit_summary.get("enrich_schema_columns") == list(
        ENRICH_REQUIRED_COLUMNS
    )

    with pytest.raises(FileNotFoundError):
        load_prepared(tmp_path / "missing.parquet")

    # Schema-stale parquet: drop a required enrich column then reload
    stale = pd.read_parquet(processed / "clean_transactions.parquet")
    stale = stale.drop(columns=["net_revenue_line"])
    stale.to_parquet(processed / "clean_transactions.parquet", index=False)
    with pytest.raises(StaleCleanCacheError, match="enrich"):
        load_prepared()

    # Restore a valid parquet so later steps in this test can use prepare again
    result.clean.to_parquet(processed / "clean_transactions.parquet", index=False)

    # Parquet exists but prep_summary missing
    (audit / "prep_summary.json").unlink()
    with pytest.raises(FileNotFoundError):
        load_prepared()

    # Excel row-limit refusal
    monkeypatch.setattr(dp, "EXCEL_MAX_DATA_ROWS", 2)
    with pytest.raises(RuntimeError, match="exceeding Excel limit"):
        prepare_dataset(path)


def test_prepare_dataset_unequal_dedupe_warning(tmp_path, monkeypatch, caplog):
    import logging

    from src import data_processing as dp
    from src import config as cfg
    from src import logging_utils as lu

    raw_dir = tmp_path / "raw"
    processed = tmp_path / "processed"
    audit = tmp_path / "audit"
    logs = tmp_path / "logs"
    for p in (raw_dir, processed, audit, logs, tmp_path / "out", tmp_path / "out" / "png"):
        p.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(dp, "DATA_RAW", raw_dir)
    monkeypatch.setattr(dp, "DATA_PROCESSED", processed)
    monkeypatch.setattr(dp, "AUDIT_DIR", audit)
    monkeypatch.setattr(cfg, "DATA_RAW", raw_dir)
    monkeypatch.setattr(cfg, "DATA_PROCESSED", processed)
    monkeypatch.setattr(cfg, "AUDIT_DIR", audit)
    monkeypatch.setattr(cfg, "LOG_DIR", logs)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", tmp_path / "out")
    monkeypatch.setattr(cfg, "PNG_DIR", tmp_path / "out" / "png")
    monkeypatch.setattr(lu, "LOG_DIR", logs)

    rows = []
    for i, sheet in enumerate(["Year 2009-2010", "Year 2009-2010", "Year 2010-2011"]):
        rows.append(
            {
                "Invoice": "SAME",
                "StockCode": "P",
                "Description": "X",
                "Quantity": 1,
                "InvoiceDate": "2011-01-01",
                "Price": 1.0,
                "Customer ID": 1.0,
                "Country": "EIRE",
                "source_sheet": sheet,
            }
        )
    # also need some unique later-sheet rows so clean is non-empty after drop
    rows.append(
        {
            "Invoice": "UNIQUE",
            "StockCode": "Q",
            "Description": "Y",
            "Quantity": 1,
            "InvoiceDate": "2011-02-01",
            "Price": 1.0,
            "Customer ID": 2.0,
            "Country": "EIRE",
            "source_sheet": "Year 2010-2011",
        }
    )
    frame = pd.DataFrame(rows)
    path = raw_dir / "online_retail_II.xlsx"
    cols = [
        "Invoice",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "Price",
        "Customer ID",
        "Country",
    ]
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for sheet, g in frame.groupby("source_sheet"):
            g[cols].to_excel(writer, sheet_name=sheet, index=False)

    with caplog.at_level(logging.WARNING, logger="retail_assessment"):
        result = prepare_dataset(path)
    assert result.audit_summary["dedupe"]["matches_multiset_max"] is False
    assert any("unequal" in r.message.lower() for r in caplog.records)


def test_prepare_dataset_raises_when_enrich_schema_invalid(
    tmp_path, monkeypatch, raw_transactions
):
    """Post-enrich schema check must fail loudly (no silent bad parquet)."""
    from src import config as cfg
    from src import data_processing as dp
    from src import logging_utils as lu

    raw_dir = tmp_path / "raw"
    processed = tmp_path / "processed"
    audit = tmp_path / "audit"
    logs = tmp_path / "logs"
    for p in (raw_dir, processed, audit, logs):
        p.mkdir()
    monkeypatch.setattr(dp, "DATA_RAW", raw_dir)
    monkeypatch.setattr(dp, "DATA_PROCESSED", processed)
    monkeypatch.setattr(dp, "AUDIT_DIR", audit)
    monkeypatch.setattr(cfg, "DATA_RAW", raw_dir)
    monkeypatch.setattr(cfg, "DATA_PROCESSED", processed)
    monkeypatch.setattr(cfg, "AUDIT_DIR", audit)
    monkeypatch.setattr(cfg, "LOG_DIR", logs)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", tmp_path / "out")
    monkeypatch.setattr(cfg, "PNG_DIR", tmp_path / "out" / "png")
    monkeypatch.setattr(lu, "LOG_DIR", logs)
    (tmp_path / "out").mkdir()
    (tmp_path / "out" / "png").mkdir()

    path = raw_dir / "online_retail_II.xlsx"
    cols = [
        "Invoice",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "Price",
        "Customer ID",
        "Country",
    ]
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for sheet, g in raw_transactions.groupby("source_sheet"):
            g[cols].to_excel(writer, sheet_name=sheet, index=False)

    monkeypatch.setattr(
        dp, "validate_enrich_schema", lambda _df: ["missing enrich column: fake"]
    )
    with pytest.raises(RuntimeError, match="missing enrich column"):
        prepare_dataset(path)


def test_description_mode_handles_all_empty_descriptions():
    from src.data_processing import _description_mode_by_stock

    df = pd.DataFrame(
        {
            "StockCode": ["A", "A"],
            "Description": [pd.NA, ""],
        }
    )
    mode = _description_mode_by_stock(df)
    assert mode.empty


def test_run_io_tasks_writes_independently(tmp_path):
    from src.data_processing import _run_io_tasks

    a = tmp_path / "a.txt"
    b = tmp_path / "b.txt"
    _run_io_tasks(
        [
            lambda: a.write_text("A", encoding="utf-8"),
            lambda: b.write_text("B", encoding="utf-8"),
        ]
    )
    assert a.read_text(encoding="utf-8") == "A"
    assert b.read_text(encoding="utf-8") == "B"

    def boom() -> None:
        raise RuntimeError("io-boom")

    with pytest.raises(RuntimeError, match="io-boom"):
        _run_io_tasks([boom])

