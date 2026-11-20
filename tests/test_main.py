"""Tests for main.py CLI orchestration (heavy deps stubbed)."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd
import pytest


@pytest.fixture
def isolated_dirs(tmp_path, monkeypatch):
    """Point config + consumers at tmp output/data/log dirs."""
    import main as main_mod
    from src import config
    from src import logging_utils as lu
    from src import email_automation as ea
    from src import packaging as pkg
    from src import validation as val
    from src import data_processing as dp
    from src import analysis as an
    from src import excel_automation as excel

    raw = tmp_path / "raw"
    processed = tmp_path / "processed"
    out = tmp_path / "output"
    png = out / "png"
    audit = out / "audit"
    logs = out / "logs"
    for p in (raw, processed, out, png, audit, logs):
        p.mkdir(parents=True, exist_ok=True)

    for mod in (config, main_mod, ea, pkg, val, dp, an, excel, lu):
        if hasattr(mod, "OUTPUT_DIR"):
            monkeypatch.setattr(mod, "OUTPUT_DIR", out)
        if hasattr(mod, "PNG_DIR"):
            monkeypatch.setattr(mod, "PNG_DIR", png)
        if hasattr(mod, "AUDIT_DIR"):
            monkeypatch.setattr(mod, "AUDIT_DIR", audit)
        if hasattr(mod, "LOG_DIR"):
            monkeypatch.setattr(mod, "LOG_DIR", logs)
        if hasattr(mod, "DATA_RAW"):
            monkeypatch.setattr(mod, "DATA_RAW", raw)
        if hasattr(mod, "DATA_PROCESSED"):
            monkeypatch.setattr(mod, "DATA_PROCESSED", processed)

    # Reset logger handlers bound to old LOG_DIR
    import logging

    logger = logging.getLogger("retail_assessment")
    for h in list(logger.handlers):
        logger.removeHandler(h)
        h.close()

    return {
        "raw": raw,
        "processed": processed,
        "out": out,
        "png": png,
        "audit": audit,
        "logs": logs,
    }


def _fake_prep(clean_transactions):
    from src.data_processing import PrepResult

    return PrepResult(
        clean=clean_transactions,
        audit_summary={
            "raw_rows": len(clean_transactions) + 1,
            "clean_rows": len(clean_transactions),
            "removed_cross_sheet_clones": 1,
            "input_sha256": "abc123",
            "dedupe": {"matches_multiset_max": True},
        },
        input_path="fake.xlsx",
        input_sha256="abc123",
    )


def test_main_happy_skip_excel_png_zip(
    isolated_dirs, clean_transactions, analysis_bundle, monkeypatch
):
    import main as main_mod

    prep = _fake_prep(clean_transactions)
    monkeypatch.setattr(main_mod, "prepare_dataset", lambda *_a, **_k: prep)
    monkeypatch.setattr(main_mod, "run_analyses", lambda *_a, **_k: analysis_bundle)
    monkeypatch.setattr(main_mod, "run_validations", lambda *a, **k: [])
    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--skip-excel", "--skip-png", "--skip-zip"],
    )
    assert main_mod.main() == 0


def test_main_validation_errors_exit_1(
    isolated_dirs, clean_transactions, analysis_bundle, monkeypatch
):
    import main as main_mod

    prep = _fake_prep(clean_transactions)
    monkeypatch.setattr(main_mod, "prepare_dataset", lambda *_a, **_k: prep)
    monkeypatch.setattr(main_mod, "run_analyses", lambda *_a, **_k: analysis_bundle)
    monkeypatch.setattr(main_mod, "run_validations", lambda *a, **k: ["bad"])
    monkeypatch.setattr("sys.argv", ["main.py", "--skip-excel", "--skip-png", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_reuse_clean_sha_mismatch(
    isolated_dirs, clean_transactions, monkeypatch
):
    import main as main_mod

    prep = _fake_prep(clean_transactions)
    prep.input_sha256 = "old"
    monkeypatch.setattr(main_mod, "load_prepared", lambda: prep)
    monkeypatch.setattr(main_mod, "sha256_file", lambda _p: "new")
    monkeypatch.setattr(main_mod, "find_input_workbook", lambda _p=None: Path("x.xlsx"))
    monkeypatch.setattr("sys.argv", ["main.py", "--reuse-clean", "--skip-excel", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_reuse_clean_schema_mismatch(
    isolated_dirs, clean_transactions, monkeypatch
):
    """--reuse-clean must fail when cached columns no longer match enrich()."""
    import main as main_mod
    from src.data_processing import StaleCleanCacheError

    def boom():
        raise StaleCleanCacheError(
            "Cached clean parquet is missing columns expected by the current enrich() "
            "schema: ['net_revenue_line']. Re-run without --reuse-clean."
        )

    monkeypatch.setattr(main_mod, "load_prepared", boom)
    monkeypatch.setattr("sys.argv", ["main.py", "--reuse-clean", "--skip-excel", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_reuse_clean_schema_mismatch_via_validate(
    isolated_dirs, clean_transactions, monkeypatch
):
    """Stubbed load_prepared still fails when validate_enrich_schema rejects columns."""
    import main as main_mod

    prep = _fake_prep(clean_transactions.drop(columns=["gross_revenue_product"], errors="ignore"))
    # Ensure the frame is actually missing a required column
    if "gross_revenue_product" in prep.clean.columns:
        prep.clean = prep.clean.drop(columns=["gross_revenue_product"])
    monkeypatch.setattr(main_mod, "load_prepared", lambda: prep)
    monkeypatch.setattr(main_mod, "sha256_file", lambda _p: prep.input_sha256)
    monkeypatch.setattr(main_mod, "find_input_workbook", lambda _p=None: Path("x.xlsx"))
    monkeypatch.setattr("sys.argv", ["main.py", "--reuse-clean", "--skip-excel", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_email_preview_only_missing_artifacts(isolated_dirs, monkeypatch):
    import main as main_mod

    monkeypatch.setattr("sys.argv", ["main.py", "--email-preview-only", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_email_preview_only_invalid_pngs(isolated_dirs, monkeypatch):
    import main as main_mod

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 2000)
    (dirs["out"] / "analysis_summary.json").write_text('{"net_revenue": 1}', encoding="utf-8")
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: ["missing PNG"])
    monkeypatch.setattr("sys.argv", ["main.py", "--email-preview-only", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_email_preview_only_manifest_mismatch(isolated_dirs, monkeypatch):
    import main as main_mod
    from PIL import Image
    from src.validation import REQUIRED_PNGS

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 2000)
    (dirs["out"] / "analysis_summary.json").write_text('{"net_revenue": 1}', encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(10, 20, 30)).save(dirs["png"] / name)
    (dirs["out"] / "office_build.json").write_text('["wrong"]', encoding="utf-8")
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr("sys.argv", ["main.py", "--email-preview-only", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_email_preview_only_success(isolated_dirs, monkeypatch):
    import main as main_mod
    from PIL import Image
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 2000)
    summary = {"net_revenue": 1.0, "email_total_revenue": 1.0, "distinct_invoices": 1}
    summary_path = dirs["out"] / "analysis_summary.json"
    summary_path.write_text(json.dumps(summary), encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(10, 20, 30)).save(dirs["png"] / name)
    artifact_paths = [wb, summary_path, *(dirs["png"] / n for n in REQUIRED_PNGS)]
    (dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in artifact_paths]), encoding="utf-8"
    )

    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr(
        main_mod,
        "write_dry_run_preview",
        lambda *a, **k: dirs["out"] / "preview.html",
    )
    (dirs["out"] / "preview.html").write_text("ok", encoding="utf-8")
    monkeypatch.setattr("sys.argv", ["main.py", "--email-preview-only", "--skip-zip"])
    assert main_mod.main() == 0


def test_main_send_no_recipient(isolated_dirs, clean_transactions, analysis_bundle, monkeypatch):
    import main as main_mod

    prep = _fake_prep(clean_transactions)
    monkeypatch.setattr(main_mod, "prepare_dataset", lambda *_a, **_k: prep)
    monkeypatch.setattr(main_mod, "run_analyses", lambda *_a, **_k: analysis_bundle)
    monkeypatch.setattr(main_mod, "run_validations", lambda *a, **k: [])

    wb = isolated_dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 100)

    def fake_build(*_a, **_k):
        return wb

    monkeypatch.setattr(main_mod, "build_workbook", fake_build)
    monkeypatch.setattr(main_mod, "build_submission_zip", lambda **k: isolated_dirs["out"] / "z.zip")
    (isolated_dirs["out"] / "z.zip").write_bytes(b"zip")
    monkeypatch.setattr(
        main_mod,
        "write_dry_run_preview",
        lambda *a, **k: isolated_dirs["out"] / "p.html",
    )
    (isolated_dirs["out"] / "p.html").write_text("p", encoding="utf-8")
    monkeypatch.setattr(main_mod, "poc_test_emails", lambda: [])
    monkeypatch.setattr("sys.argv", ["main.py", "--send", "--skip-png"])
    # skip-png but not skip-excel — will try email after build
    # Actually with skip-png and not skip-excel, after validate it continues if both skip_excel or skip_png:
    # `if args.skip_excel or args.skip_png: return 0` — so --send never reached!
    # Need neither skip for send path after full pipeline, OR email-preview-only + send

    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--email-preview-only", "--send", "--skip-zip"],
    )
    # Need artifacts for preview-only
    from PIL import Image
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    summary_path = isolated_dirs["out"] / "analysis_summary.json"
    summary_path.write_text(
        json.dumps({"net_revenue": 1, "email_total_revenue": 1}), encoding="utf-8"
    )
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(1, 2, 3)).save(isolated_dirs["png"] / name)
    arts = [wb, summary_path, *(isolated_dirs["png"] / n for n in REQUIRED_PNGS)]
    (isolated_dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in arts]), encoding="utf-8"
    )
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    assert main_mod.main() == 2


def test_main_send_workbook_too_large(isolated_dirs, monkeypatch):
    """With --skip-zip and an oversized workbook, send must hard-fail (no alternate ZIP)."""
    import main as main_mod
    from PIL import Image
    from src.config import ATTACHMENT_LIMIT_BYTES
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"0" * (ATTACHMENT_LIMIT_BYTES + 1))
    summary_path = dirs["out"] / "analysis_summary.json"
    summary_path.write_text(json.dumps({"net_revenue": 1}), encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(1, 2, 3)).save(dirs["png"] / name)
    arts = [wb, summary_path, *(dirs["png"] / n for n in REQUIRED_PNGS)]
    (dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in arts]), encoding="utf-8"
    )
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr(
        main_mod, "write_dry_run_preview", lambda *a, **k: dirs["out"] / "p.html"
    )
    (dirs["out"] / "p.html").write_text("p", encoding="utf-8")
    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--email-preview-only", "--send", "--send-to", "a@b.com", "--skip-zip"],
    )
    assert main_mod.main() == 2


def test_main_send_success_and_failure(isolated_dirs, monkeypatch):
    import main as main_mod
    from PIL import Image
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 1000)
    summary_path = dirs["out"] / "analysis_summary.json"
    summary_path.write_text(json.dumps({"net_revenue": 1}), encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(1, 2, 3)).save(dirs["png"] / name)
    arts = [wb, summary_path, *(dirs["png"] / n for n in REQUIRED_PNGS)]
    (dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in arts]), encoding="utf-8"
    )
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr(
        main_mod, "write_dry_run_preview", lambda *a, **k: dirs["out"] / "p.html"
    )
    (dirs["out"] / "p.html").write_text("p", encoding="utf-8")

    monkeypatch.setattr(
        main_mod,
        "send_submission_email",
        lambda *a, **k: {"status": "sent_api_returned"},
    )
    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--email-preview-only", "--send", "--send-to", "me@test.com", "--skip-zip"],
    )
    assert main_mod.main() == 0

    monkeypatch.setattr(
        main_mod,
        "send_submission_email",
        lambda *a, **k: {"status": "failed", "error": "boom"},
    )
    assert main_mod.main() == 3


def test_main_send_uses_poc_test_email(isolated_dirs, monkeypatch):
    import main as main_mod
    from PIL import Image
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 1000)
    summary_path = dirs["out"] / "analysis_summary.json"
    summary_path.write_text(json.dumps({"net_revenue": 1}), encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(1, 2, 3)).save(dirs["png"] / name)
    arts = [wb, summary_path, *(dirs["png"] / n for n in REQUIRED_PNGS)]
    (dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in arts]), encoding="utf-8"
    )
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr(
        main_mod, "write_dry_run_preview", lambda *a, **k: dirs["out"] / "p.html"
    )
    (dirs["out"] / "p.html").write_text("p", encoding="utf-8")
    monkeypatch.setattr(main_mod, "poc_test_emails", lambda: ["poc@test.com"])
    called = {}

    def capture_send(summary, workbook, recipients, **kwargs):
        called["recipients"] = recipients
        called["kwargs"] = kwargs
        return {"status": "sent_api_returned"}

    monkeypatch.setattr(main_mod, "send_submission_email", capture_send)
    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--email-preview-only", "--send", "--skip-zip"],
    )
    assert main_mod.main() == 0
    assert called["recipients"] == ["poc@test.com"]
    assert called["kwargs"].get("zip_path") is None
    assert called["kwargs"].get("attach_zip") is False


def test_main_send_attaches_email_zip(isolated_dirs, monkeypatch):
    """--send must reuse the packaged email_zip (not zip_path=None / attach_zip=False)."""
    import main as main_mod
    from PIL import Image
    from src.config import ATTACHMENT_LIMIT_BYTES
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 1000)
    summary_path = dirs["out"] / "analysis_summary.json"
    summary_path.write_text(json.dumps({"net_revenue": 1}), encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(1, 2, 3)).save(dirs["png"] / name)
    arts = [wb, summary_path, *(dirs["png"] / n for n in REQUIRED_PNGS)]
    (dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in arts]), encoding="utf-8"
    )

    full_zip = dirs["out"] / "full.zip"
    full_zip.write_bytes(b"z" * 200)
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr(main_mod, "build_submission_zip", lambda **k: full_zip)
    monkeypatch.setattr(
        main_mod, "write_dry_run_preview", lambda *a, **k: dirs["out"] / "p.html"
    )
    (dirs["out"] / "p.html").write_text("p", encoding="utf-8")

    called = {}

    def capture_send(summary, workbook, recipients, **kwargs):
        called["kwargs"] = kwargs
        return {"status": "sent_api_returned"}

    monkeypatch.setattr(main_mod, "send_submission_email", capture_send)
    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--email-preview-only", "--send", "--send-to", "me@test.com"],
    )
    assert main_mod.main() == 0
    assert called["kwargs"]["zip_path"] == full_zip
    assert called["kwargs"]["attach_zip"] is True
    assert called["kwargs"]["attach_workbook"] is True
    assert full_zip.stat().st_size <= ATTACHMENT_LIMIT_BYTES


def test_main_send_large_workbook_uses_email_safe_zip(isolated_dirs, monkeypatch):
    """Oversized workbook + email-safe ZIP should still send (ZIP only)."""
    import main as main_mod
    from PIL import Image
    from src.config import ATTACHMENT_LIMIT_BYTES
    from src.validation import REQUIRED_PNGS
    from src.data_processing import sha256_file

    dirs = isolated_dirs
    wb = dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"0" * (ATTACHMENT_LIMIT_BYTES + 1))
    summary_path = dirs["out"] / "analysis_summary.json"
    summary_path.write_text(json.dumps({"net_revenue": 1}), encoding="utf-8")
    for name in REQUIRED_PNGS:
        Image.new("RGB", (30, 20), color=(1, 2, 3)).save(dirs["png"] / name)
    arts = [wb, summary_path, *(dirs["png"] / n for n in REQUIRED_PNGS)]
    (dirs["out"] / "office_build.json").write_text(
        json.dumps([sha256_file(p) for p in arts]), encoding="utf-8"
    )

    big_zip = dirs["out"] / "big.zip"
    big_zip.write_bytes(b"z" * (ATTACHMENT_LIMIT_BYTES + 1))
    safe = dirs["out"] / "safe.zip"
    safe.write_bytes(b"safe")
    monkeypatch.setattr(main_mod, "validate_pngs", lambda: [])
    monkeypatch.setattr(main_mod, "build_submission_zip", lambda **k: big_zip)
    monkeypatch.setattr(main_mod, "build_email_safe_zip", lambda *a, **k: safe)
    monkeypatch.setattr(
        main_mod, "write_dry_run_preview", lambda *a, **k: dirs["out"] / "p.html"
    )
    (dirs["out"] / "p.html").write_text("p", encoding="utf-8")

    called = {}

    def capture_send(summary, workbook, recipients, **kwargs):
        called["kwargs"] = kwargs
        return {"status": "sent_api_returned"}

    monkeypatch.setattr(main_mod, "send_submission_email", capture_send)
    monkeypatch.setattr(
        "sys.argv",
        ["main.py", "--email-preview-only", "--send", "--send-to", "a@b.com"],
    )
    assert main_mod.main() == 0
    assert called["kwargs"]["zip_path"] == safe
    assert called["kwargs"]["attach_zip"] is True
    assert called["kwargs"]["attach_workbook"] is False


def test_main_full_pipeline_with_excel_stub(
    isolated_dirs, clean_transactions, analysis_bundle, monkeypatch
):
    import main as main_mod
    from PIL import Image
    from src.validation import REQUIRED_PNGS

    prep = _fake_prep(clean_transactions)
    monkeypatch.setattr(main_mod, "prepare_dataset", lambda *_a, **_k: prep)
    monkeypatch.setattr(main_mod, "run_analyses", lambda *_a, **_k: analysis_bundle)
    monkeypatch.setattr(main_mod, "run_validations", lambda *a, **k: [])

    wb = isolated_dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 5000)
    for name in REQUIRED_PNGS:
        Image.new("RGB", (40, 20), color=(5, 6, 7)).save(isolated_dirs["png"] / name)

    monkeypatch.setattr(main_mod, "build_workbook", lambda *a, **k: wb)
    zip_path = isolated_dirs["out"] / "full.zip"
    zip_path.write_bytes(b"z" * 100)
    monkeypatch.setattr(main_mod, "build_submission_zip", lambda **k: zip_path)
    monkeypatch.setattr(
        main_mod,
        "write_dry_run_preview",
        lambda *a, **k: isolated_dirs["out"] / "p.html",
    )
    (isolated_dirs["out"] / "p.html").write_text("p", encoding="utf-8")
    monkeypatch.setattr("sys.argv", ["main.py"])
    assert main_mod.main() == 0
    assert (isolated_dirs["out"] / "office_build.json").exists()
    assert (isolated_dirs["out"] / "analysis_summary.json").exists()


def test_main_unhandled_exception(isolated_dirs, monkeypatch):
    import main as main_mod

    monkeypatch.setattr(
        main_mod,
        "prepare_dataset",
        lambda *_a, **_k: (_ for _ in ()).throw(RuntimeError("explode")),
    )
    monkeypatch.setattr("sys.argv", ["main.py", "--skip-excel", "--skip-zip"])
    assert main_mod.main() == 1


def test_main_large_zip_builds_email_safe(
    isolated_dirs, clean_transactions, analysis_bundle, monkeypatch
):
    import main as main_mod
    from PIL import Image
    from src.validation import REQUIRED_PNGS

    prep = _fake_prep(clean_transactions)
    monkeypatch.setattr(main_mod, "prepare_dataset", lambda *_a, **_k: prep)
    monkeypatch.setattr(main_mod, "run_analyses", lambda *_a, **_k: analysis_bundle)
    monkeypatch.setattr(main_mod, "run_validations", lambda *a, **k: [])
    wb = isolated_dirs["out"] / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 18_000_001)
    for name in REQUIRED_PNGS:
        Image.new("RGB", (40, 20), color=(5, 6, 7)).save(isolated_dirs["png"] / name)
    monkeypatch.setattr(main_mod, "build_workbook", lambda *a, **k: wb)
    big_zip = isolated_dirs["out"] / "big.zip"
    big_zip.write_bytes(b"z" * 18_000_001)
    safe = isolated_dirs["out"] / "safe.zip"
    safe.write_bytes(b"safe")
    monkeypatch.setattr(main_mod, "build_submission_zip", lambda **k: big_zip)
    monkeypatch.setattr(main_mod, "build_email_safe_zip", lambda *a, **k: safe)
    monkeypatch.setattr(
        main_mod,
        "write_dry_run_preview",
        lambda *a, **k: isolated_dirs["out"] / "p.html",
    )
    (isolated_dirs["out"] / "p.html").write_text("p", encoding="utf-8")
    # skip_excel or skip_png early return — need full path without early return
    # With neither skip, after validate continues to package
    monkeypatch.setattr("sys.argv", ["main.py"])
    assert main_mod.main() == 0
