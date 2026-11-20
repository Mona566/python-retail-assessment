"""Tests for ZIP packaging helpers."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from PIL import Image

from src.packaging import build_email_safe_zip, build_submission_zip


def test_email_safe_zip_contains_source_and_pngs_not_workbook(tmp_path, monkeypatch):
    from src import packaging as pkg
    from src import config as cfg

    # Isolate output under tmp
    monkeypatch.setattr(pkg, "OUTPUT_DIR", tmp_path)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", tmp_path)

    png_dir = tmp_path / "png"
    audit_dir = tmp_path / "audit"
    png_dir.mkdir()
    audit_dir.mkdir()
    for name in [
        "Pivot_Monthly_Performance.png",
        "Pivot_Country_Performance.png",
    ]:
        Image.new("RGB", (20, 10), color=(240, 240, 240)).save(png_dir / name)
    (audit_dir / "prep_summary.json").write_text('{"ok": true}', encoding="utf-8")
    summary = tmp_path / "analysis_summary.json"
    summary.write_text('{"net_revenue": 1}', encoding="utf-8")

    # candidate_name from env / default
    zip_path = build_email_safe_zip(png_dir, audit_dir, summary)
    assert zip_path.exists()
    assert zip_path.stat().st_size > 0

    with zipfile.ZipFile(zip_path) as zf:
        names = set(zf.namelist())
    assert "main.py" in names
    assert "README.md" in names
    assert "docs/ASSIGNMENT_WALKTHROUGH.md" in names
    assert "tests/test_analysis.py" in names
    assert "manifest.json" in names
    assert "ATTACHMENT_NOTE.txt" in names
    assert any(n.startswith("png/") for n in names)
    assert "Retail_Analysis.xlsx" not in names

    # Rebuild when zip already exists
    again = build_email_safe_zip(png_dir, audit_dir, summary)
    assert again == zip_path


def test_build_submission_zip_includes_workbook(tmp_path, monkeypatch):
    from src import packaging as pkg
    from src import config as cfg

    monkeypatch.setattr(pkg, "OUTPUT_DIR", tmp_path)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", tmp_path)
    monkeypatch.setenv("CANDIDATE_NAME", "Test User")

    png_dir = tmp_path / "png"
    audit_dir = tmp_path / "audit"
    png_dir.mkdir()
    audit_dir.mkdir()
    Image.new("RGB", (20, 10)).save(png_dir / "Pivot_Monthly_Performance.png")
    (audit_dir / "prep_summary.json").write_text("{}", encoding="utf-8")
    summary = tmp_path / "analysis_summary.json"
    summary.write_text("{}", encoding="utf-8")
    workbook = tmp_path / "Retail_Analysis.xlsx"
    workbook.write_bytes(b"workbook-content")

    with pytest.raises(FileNotFoundError):
        build_submission_zip(tmp_path / "missing.xlsx", png_dir, audit_dir, summary)

    zip_path = build_submission_zip(workbook, png_dir, audit_dir, summary)
    assert zip_path.name == "TestUser_PythonAssessment.zip"
    with zipfile.ZipFile(zip_path) as zf:
        names = set(zf.namelist())
    assert "Retail_Analysis.xlsx" in names
    assert "manifest.json" in names

    # Rebuild when zip exists (unlink path)
    zip_path.write_bytes(b"stale")
    rebuilt = build_submission_zip(workbook, png_dir, audit_dir, summary)
    assert rebuilt.exists()
    with zipfile.ZipFile(rebuilt) as zf:
        assert "Retail_Analysis.xlsx" in zf.namelist()
