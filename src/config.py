"""Paths and environment-backed settings (no hard-coded secrets)."""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

from src.constants import Artifacts, EmailConst

PROJECT_ROOT = Path(__file__).resolve().parents[1]
load_dotenv(PROJECT_ROOT / ".env")

DATA_RAW = PROJECT_ROOT / "data" / "raw"
DATA_PROCESSED = PROJECT_ROOT / "data" / "processed"
OUTPUT_DIR = PROJECT_ROOT / "output"
PNG_DIR = OUTPUT_DIR / "png"
AUDIT_DIR = OUTPUT_DIR / "audit"
LOG_DIR = OUTPUT_DIR / "logs"
POC_DIR = OUTPUT_DIR / "poc"

INPUT_XLSX_NAME = "online_retail_II.xlsx"
OUTPUT_XLSX_NAME = "Retail_Analysis.xlsx"

# Typical mailbox attachment ceiling used for workbook/ZIP send decisions.
ATTACHMENT_LIMIT_BYTES = 18_000_000
ATTACHMENT_LIMIT_MB = ATTACHMENT_LIMIT_BYTES / 1_000_000  # 18.0 — compare float MB sizes

# Artifact basename re-exports for path construction across modules.
ANALYSIS_SUMMARY_JSON = Artifacts.ANALYSIS_SUMMARY_JSON
OFFICE_BUILD_JSON = Artifacts.OFFICE_BUILD_JSON
CLEAN_PARQUET_NAME = Artifacts.CLEAN_PARQUET
PREP_SUMMARY_JSON = Artifacts.PREP_SUMMARY_JSON
EMAIL_PREVIEW_DIRNAME = EmailConst.PREVIEW_DIRNAME


def candidate_name() -> str:
    """Display name for ZIP filename and email subject (from CANDIDATE_NAME)."""
    return os.getenv("CANDIDATE_NAME", "Candidate").strip() or "Candidate"


def submission_email() -> str | None:
    """Official submission inbox from SUBMISSION_EMAIL (never auto-used as --send target)."""
    value = os.getenv("SUBMISSION_EMAIL", "").strip()
    return value or None


def poc_test_email() -> str | None:
    """First PoC recipient (backward compatible)."""
    emails = poc_test_emails()
    return emails[0] if emails else None


def poc_test_emails() -> list[str]:
    """Comma/semicolon-separated PoC recipients from POC_TEST_EMAIL."""
    value = os.getenv("POC_TEST_EMAIL", "").strip()
    if not value:
        return []
    parts = [p.strip() for p in value.replace(";", ",").split(",")]
    return [p for p in parts if p]


def ensure_dirs() -> None:
    """Create data/output/audit/log directories if missing."""
    for path in (
        DATA_RAW,
        DATA_PROCESSED,
        OUTPUT_DIR,
        PNG_DIR,
        AUDIT_DIR,
        LOG_DIR,
        POC_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)
