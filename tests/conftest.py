"""Shared fixtures: tiny synthetic Online Retail II–like frames (no Excel/Outlook)."""

from __future__ import annotations

import pandas as pd
import pytest

# Repo root is on sys.path via pytest.ini ``pythonpath = .`` (see path_setup.py for CLI).
from src.data_processing import enrich


def _row(
    *,
    sheet: str,
    row: int,
    invoice: str,
    stock: str,
    desc: str,
    qty: float,
    price: float,
    date: str,
    customer,
    country: str,
    load_id: int,
) -> dict:
    return {
        "load_row_id": load_id,
        "source_sheet": sheet,
        "source_row": row,
        "Invoice": invoice,
        "StockCode": stock,
        "Description": desc,
        "Quantity": qty,
        "InvoiceDate": date,
        "Price": price,
        "Customer ID": customer,
        "Country": country,
    }


@pytest.fixture
def raw_transactions() -> pd.DataFrame:
    """
    Minimal two-sheet raw extract covering:
    - cross-sheet clone (same business key on both years)
    - within-sheet exact duplicate (kept + flagged)
    - UK + non-UK sales
    - cancellation (C…)
    - non-cancellation return (negative qty)
    - missing Customer ID
    - two products with enough volume for return-rate ranking
    """
    rows = [
        # Shared business key on both sheets → earlier sheet clone removed
        _row(
            sheet="Year 2009-2010",
            row=2,
            invoice="10001",
            stock="P1",
            desc="MUG",
            qty=10,
            price=2.0,
            date="2010-06-01 10:00:00",
            customer=1000.0,
            country="United Kingdom",
            load_id=1,
        ),
        _row(
            sheet="Year 2010-2011",
            row=2,
            invoice="10001",
            stock="P1",
            desc="MUG",
            qty=10,
            price=2.0,
            date="2010-06-01 10:00:00",
            customer=1000.0,
            country="United Kingdom",
            load_id=2,
        ),
        # Within-sheet duplicate (same key twice on 2010-2011) — keep both
        _row(
            sheet="Year 2010-2011",
            row=3,
            invoice="10002",
            stock="P1",
            desc="MUG",
            qty=5,
            price=2.0,
            date="2011-03-01 11:00:00",
            customer=1000.0,
            country="United Kingdom",
            load_id=3,
        ),
        _row(
            sheet="Year 2010-2011",
            row=4,
            invoice="10002",
            stock="P1",
            desc="MUG",
            qty=5,
            price=2.0,
            date="2011-03-01 11:00:00",
            customer=1000.0,
            country="United Kingdom",
            load_id=4,
        ),
        # Non-UK sales (EIRE) — complete March + April 2011 coverage helpers
        _row(
            sheet="Year 2010-2011",
            row=5,
            invoice="20001",
            stock="P2",
            desc="BOWL",
            qty=50,
            price=3.0,
            date="2011-03-01 09:00:00",
            customer=2000.0,
            country="EIRE",
            load_id=5,
        ),
        _row(
            sheet="Year 2010-2011",
            row=6,
            invoice="20002",
            stock="P2",
            desc="BOWL",
            qty=50,
            price=3.0,
            date="2011-03-31 15:00:00",
            customer=2000.0,
            country="EIRE",
            load_id=6,
        ),
        _row(
            sheet="Year 2010-2011",
            row=7,
            invoice="20003",
            stock="P2",
            desc="BOWL",
            qty=40,
            price=3.0,
            date="2011-04-01 12:00:00",
            customer=2000.0,
            country="EIRE",
            load_id=7,
        ),
        _row(
            sheet="Year 2010-2011",
            row=8,
            invoice="20004",
            stock="P2",
            desc="BOWL",
            qty=40,
            price=3.0,
            date="2011-04-30 16:00:00",
            customer=None,
            country="EIRE",
            load_id=8,
        ),
        # Cancellation
        _row(
            sheet="Year 2010-2011",
            row=9,
            invoice="C30001",
            stock="P1",
            desc="MUG",
            qty=-2,
            price=2.0,
            date="2011-04-15 10:00:00",
            customer=1000.0,
            country="United Kingdom",
            load_id=9,
        ),
        # Non-cancellation return (Price=0 adjustment style)
        _row(
            sheet="Year 2010-2011",
            row=10,
            invoice="30002",
            stock="P2",
            desc="BOWL",
            qty=-20,
            price=0.0,
            date="2011-04-20 10:00:00",
            customer=2000.0,
            country="EIRE",
            load_id=10,
        ),
        # Extra P1 sales so P1 has volume for product ranking
        _row(
            sheet="Year 2010-2011",
            row=11,
            invoice="10003",
            stock="P1",
            desc="MUG",
            qty=80,
            price=2.0,
            date="2011-04-10 14:00:00",
            customer=1000.0,
            country="France",
            load_id=11,
        ),
        # Partial December edge month
        _row(
            sheet="Year 2010-2011",
            row=12,
            invoice="40001",
            stock="P1",
            desc="MUG",
            qty=1,
            price=2.0,
            date="2011-12-01 10:00:00",
            customer=1000.0,
            country="France",
            load_id=12,
        ),
        _row(
            sheet="Year 2010-2011",
            row=13,
            invoice="40002",
            stock="P1",
            desc="MUG",
            qty=1,
            price=2.0,
            date="2011-12-09 10:00:00",
            customer=1000.0,
            country="France",
            load_id=13,
        ),
    ]
    # Pad P2 to >= 50 units sold for return-rate eligibility (already 180 sold)
    return pd.DataFrame(rows)


@pytest.fixture
def clean_transactions(raw_transactions: pd.DataFrame) -> pd.DataFrame:
    from src.data_processing import apply_cross_sheet_dedupe

    kept, _removed, _dedupe = apply_cross_sheet_dedupe(raw_transactions)
    return enrich(kept)


@pytest.fixture
def analysis_bundle(clean_transactions: pd.DataFrame):
    from src.analysis import run_analyses

    return run_analyses(clean_transactions)


@pytest.fixture(autouse=True)
def isolate_analysis_audit(monkeypatch, tmp_path):
    from src import config

    monkeypatch.setattr(config, "AUDIT_DIR", tmp_path)
