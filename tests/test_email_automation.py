"""Tests for email HTML construction and send helpers (no real Outlook)."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.email_automation import (
    INLINE_PNGS,
    _collect_attachments,
    _fmt_money,
    build_email_html,
    send_submission_email,
    write_dry_run_preview,
)


def test_build_email_html_contains_required_metrics(analysis_bundle, tmp_path, monkeypatch):
    from PIL import Image
    from src import email_automation as ea

    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)
    for name, _ in INLINE_PNGS:
        Image.new("RGB", (40, 20)).save(tmp_path / name)
    summary = analysis_bundle.summary
    html = build_email_html(summary, embed_images=False, workbook_mb=219.5)
    assert "Python take-home assessment" in html or "take-home" in html.lower()
    assert "Total revenue (Net)" in html
    assert _fmt_money(summary.get("email_total_revenue")) in html
    assert _fmt_money(summary.get("net_revenue_after_cancellations")) in html
    inv = summary.get("distinct_invoices")
    cust = summary.get("distinct_customers")
    assert (f"{inv:,}" if isinstance(inv, int) else str(inv)) in html
    assert (f"{cust:,}" if isinstance(cust, int) else str(cust)) in html
    assert str(summary.get("top_product_stockcode")) in html
    assert str(summary.get("top_non_uk_country")) in html
    assert str(summary.get("best_complete_month")) in html
    assert "Distinct invoices" in html
    assert "Distinct customers" in html
    assert "Top product" in html
    assert "Top non-UK country" in html
    assert "Best complete sales month" in html
    assert "Net after cancellations" in html
    assert "Attachment note" in html or "Retail_Analysis.xlsx" in html
    for name, _cid in INLINE_PNGS:
        assert name in html
        assert f"src='{name}'" in html
    assert "data:image/png;base64," not in html


def test_build_email_html_embeds_base64_when_pngs_exist(analysis_bundle, monkeypatch, tmp_path):
    from PIL import Image
    from src import email_automation as ea

    for name, _ in INLINE_PNGS:
        img = Image.new("RGB", (40, 20), color=(255, 255, 255))
        img.save(tmp_path / name)
    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)

    html = build_email_html(analysis_bundle.summary, embed_images=True, workbook_mb=500.0)
    assert "data:image/png;base64," in html
    assert html.count("data:image/png;base64,") == 2
    assert "not attached to this email" in html
    assert "alternate delivery channel requires confirmation" in html

def test_cid_draft_has_two_inline_parts_and_workbook(analysis_bundle, monkeypatch, tmp_path):
    from email import policy
    from email.parser import BytesParser

    from PIL import Image
    from src import email_automation as ea

    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)
    for name, _ in INLINE_PNGS:
        Image.new("RGB", (40, 20)).save(tmp_path / name)
    workbook = tmp_path / "Retail_Analysis.xlsx"
    workbook.write_bytes(b"synthetic workbook bytes")
    draft = tmp_path / "submission.eml"
    ea._write_cid_draft(draft, analysis_bundle.summary, workbook)
    message = BytesParser(policy=policy.default).parsebytes(draft.read_bytes())
    html = message.get_body(preferencelist=("html",)).get_content()
    assert html.count("cid:") == 2
    assert "data:image" not in html
    inline_parts = [part for part in message.walk() if part.get_content_disposition() == "inline"]
    assert {part["Content-ID"] for part in inline_parts} == {f"<{cid}>" for _, cid in INLINE_PNGS}
    attachments = list(message.iter_attachments())
    assert attachments[0].get_filename() == "Retail_Analysis.xlsx"
    assert attachments[0].get_payload(decode=True) == workbook.read_bytes()


def test_fmt_money():
    assert _fmt_money(None) == "n/a"
    assert _fmt_money(1234.5) == "£1,234.50"


def _png_fixture(tmp_path):
    from PIL import Image

    for name, _ in INLINE_PNGS:
        Image.new("RGB", (40, 20), color=(200, 100, 50)).save(tmp_path / name)


def test_build_email_html_missing_png_raises(analysis_bundle, tmp_path, monkeypatch):
    from src import email_automation as ea

    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)
    with pytest.raises(FileNotFoundError):
        build_email_html(analysis_bundle.summary, embed_images=False, workbook_mb=1.0)


def test_build_email_html_cid_and_workbook_auto_stat(
    analysis_bundle, tmp_path, monkeypatch
):
    from src import email_automation as ea
    from src import config as cfg

    _png_fixture(tmp_path)
    monkeypatch.setattr(ea, "PNG_DIR", tmp_path)
    out = tmp_path / "out"
    out.mkdir()
    wb = out / "Retail_Analysis.xlsx"
    wb.write_bytes(b"x" * 1000)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out)
    monkeypatch.setattr(cfg, "OUTPUT_XLSX_NAME", "Retail_Analysis.xlsx")

    html = build_email_html(analysis_bundle.summary, cid_images=True, workbook_mb=None)
    assert "cid:monthly_pivot" in html
    assert "cid:country_pivot" in html


def test_write_dry_run_preview(analysis_bundle, tmp_path, monkeypatch):
    from src import email_automation as ea
    from src import config as cfg

    out = tmp_path / "output"
    png = tmp_path / "png"
    out.mkdir()
    png.mkdir()
    _png_fixture(png)
    monkeypatch.setattr(ea, "OUTPUT_DIR", out)
    monkeypatch.setattr(ea, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out)
    monkeypatch.setattr(cfg, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "AUDIT_DIR", tmp_path / "audit")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "audit").mkdir()
    (tmp_path / "logs").mkdir()

    workbook = out / "Retail_Analysis.xlsx"
    workbook.write_bytes(b"workbook-bytes")
    zip_path = out / "pack.zip"
    zip_path.write_bytes(b"zip")

    preview = write_dry_run_preview(analysis_bundle.summary, workbook, zip_path)
    assert preview.exists()
    assert (out / "email_preview" / "submission_email_send.html").exists()
    assert (out / "email_preview" / "submission_email_meta.json").exists()
    assert (out / "email_preview" / "submission_email_cid.eml").exists()


def test_collect_attachments_small_and_oversized(tmp_path, monkeypatch):
    from src import email_automation as ea
    from src import config as cfg

    png = tmp_path / "png"
    png.mkdir()
    _png_fixture(png)
    monkeypatch.setattr(ea, "PNG_DIR", png)
    monkeypatch.setattr(ea, "OUTPUT_DIR", tmp_path)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", tmp_path)
    monkeypatch.setattr(cfg, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "AUDIT_DIR", tmp_path / "audit")
    (tmp_path / "audit").mkdir()
    (tmp_path / "analysis_summary.json").write_text("{}", encoding="utf-8")

    small_wb = tmp_path / "small.xlsx"
    small_wb.write_bytes(b"x" * 100)
    small_zip = tmp_path / "small.zip"
    small_zip.write_bytes(b"z" * 100)
    atts, notes = _collect_attachments(
        small_wb, small_zip, attach_workbook=True, attach_zip=True, attach_pngs=True
    )
    assert small_wb in atts and small_zip in atts
    assert notes == []

    big_wb = tmp_path / "big.xlsx"
    big_wb.write_bytes(b"x" * 18_000_001)
    big_zip = tmp_path / "big.zip"
    big_zip.write_bytes(b"z" * 18_000_001)
    safe = tmp_path / "Candidate_PythonAssessment_EmailSafe.zip"
    safe.write_bytes(b"safe")

    with patch("src.email_automation.build_email_safe_zip", return_value=safe):
        atts2, notes2 = _collect_attachments(
            big_wb, big_zip, attach_workbook=True, attach_zip=True, attach_pngs=True
        )
    assert big_wb not in atts2
    assert safe in atts2
    assert any("Full workbook not emailed" in n for n in notes2)
    assert any("email-safe ZIP" in n for n in notes2)

    empty_png = tmp_path / "empty_png"
    empty_png.mkdir()
    monkeypatch.setattr(ea, "PNG_DIR", empty_png)
    with pytest.raises(FileNotFoundError):
        _collect_attachments(
            small_wb, None, attach_workbook=False, attach_zip=False, attach_pngs=True
        )


def test_send_submission_email_dry_run_and_send(analysis_bundle, tmp_path, monkeypatch):
    from src import email_automation as ea
    from src import config as cfg

    out = tmp_path / "output"
    png = tmp_path / "png"
    out.mkdir()
    png.mkdir()
    _png_fixture(png)
    monkeypatch.setattr(ea, "OUTPUT_DIR", out)
    monkeypatch.setattr(ea, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out)
    monkeypatch.setattr(cfg, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "AUDIT_DIR", tmp_path / "audit")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "audit").mkdir()
    (tmp_path / "logs").mkdir()

    workbook = out / "Retail_Analysis.xlsx"
    workbook.write_bytes(b"wb")

    with pytest.raises(ValueError, match="No recipients"):
        send_submission_email(analysis_bundle.summary, workbook, recipients=[])
    with pytest.raises(ValueError, match="Invalid email recipient"):
        send_submission_email(
            analysis_bundle.summary, workbook, recipients=["not-an-email"]
        )

    dry = send_submission_email(
        analysis_bundle.summary, workbook, recipients=["a@b.com"], send=False
    )
    assert dry["status"] == "dry_run"
    assert (out / "email_preview" / "submission_receipt.json").exists()

    # Mock appscript success
    fake_appscript = MagicMock()
    fake_k = MagicMock()
    fake_k.outgoing_message = "om"
    fake_k.to_recipient = "tr"
    fake_k.attachment = "att"
    fake_k.subject = "subject"
    fake_k.content = "content"
    fake_k.email_address = "email_address"
    fake_k.address = "address"
    fake_k.file = "file"
    outlook = MagicMock()
    msg = MagicMock()
    outlook.make.return_value = msg
    fake_appscript.app.return_value = outlook
    fake_mactypes = MagicMock()
    fake_mactypes.Alias = lambda p: f"alias:{p}"

    with (
        patch("src.email_automation.app", fake_appscript.app),
        patch("src.email_automation.k", fake_k),
        patch("src.email_automation.Alias", fake_mactypes.Alias),
    ):
        receipt = send_submission_email(
            analysis_bundle.summary,
            workbook,
            recipients=["a@b.com"],
            send=True,
            attach_zip=False,
        )
    assert receipt["status"] == "sent_api_returned"
    assert receipt["sent_to"] == ["a@b.com"]

    # Failure path
    outlook.make.side_effect = RuntimeError("outlook down")
    with (
        patch("src.email_automation.app", fake_appscript.app),
        patch("src.email_automation.k", fake_k),
        patch("src.email_automation.Alias", fake_mactypes.Alias),
    ):
        failed = send_submission_email(
            analysis_bundle.summary,
            workbook,
            recipients=["a@b.com"],
            send=True,
            attach_zip=False,
        )
    assert failed["status"] == "failed"
    assert "RuntimeError" in failed["error"]


def test_send_injects_attachment_notes_into_html(analysis_bundle, tmp_path, monkeypatch):
    """Oversized workbook notes must be spliced into the Outlook HTML body."""
    from src import config as cfg
    from src import email_automation as ea
    from src.config import ATTACHMENT_LIMIT_BYTES

    out = tmp_path / "output"
    png = tmp_path / "png"
    out.mkdir()
    png.mkdir()
    _png_fixture(png)
    monkeypatch.setattr(ea, "OUTPUT_DIR", out)
    monkeypatch.setattr(ea, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "OUTPUT_DIR", out)
    monkeypatch.setattr(cfg, "PNG_DIR", png)
    monkeypatch.setattr(cfg, "AUDIT_DIR", tmp_path / "audit")
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_path / "logs")
    (tmp_path / "audit").mkdir()
    (tmp_path / "logs").mkdir()

    workbook = out / "Retail_Analysis.xlsx"
    workbook.write_bytes(b"x" * (ATTACHMENT_LIMIT_BYTES + 1))

    fake_k = MagicMock()
    fake_k.outgoing_message = "om"
    fake_k.to_recipient = "tr"
    fake_k.attachment = "att"
    fake_k.subject = "subject"
    fake_k.content = "content"
    fake_k.email_address = "email_address"
    fake_k.address = "address"
    fake_k.file = "file"
    outlook = MagicMock()
    msg = MagicMock()
    outlook.make.return_value = msg
    fake_app = MagicMock(return_value=outlook)

    with (
        patch("src.email_automation.app", fake_app),
        patch("src.email_automation.k", fake_k),
        patch("src.email_automation.Alias", lambda p: f"alias:{p}"),
    ):
        receipt = send_submission_email(
            analysis_bundle.summary,
            workbook,
            recipients=["a@b.com"],
            send=True,
            attach_workbook=True,
            attach_zip=False,
        )

    assert receipt["status"] == "sent_api_returned"
    assert receipt["notes"]
    assert any("Full workbook not emailed" in n for n in receipt["notes"])
    sent_html = outlook.make.call_args.kwargs["with_properties"][fake_k.content]
    assert "Full workbook not emailed" in sent_html
