"""
Outlook submission email (macOS appscript).

Default is dry-run: write a preview HTML + receipt draft, do not send.
Use send_submission_email(..., send=True) only after explicit confirmation.

Assignment requires two Pivot PNGs embedded inline in the HTML body (not
attachments-only). On macOS Outlook we embed via HTML data-URI (visible in
Outlook compose/Sent). Gmail often strips data-URIs — we ALSO attach the same
PNGs as files so recipients still have the images if the client strips embeds.
"""

from __future__ import annotations

import base64
import json
import shutil
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from appscript import app, k
from mactypes import Alias

from src.config import (
    ANALYSIS_SUMMARY_JSON,
    ATTACHMENT_LIMIT_BYTES,
    ATTACHMENT_LIMIT_MB,
    AUDIT_DIR,
    EMAIL_PREVIEW_DIRNAME,
    OUTPUT_DIR,
    OUTPUT_XLSX_NAME,
    PNG_DIR,
    candidate_name,
    ensure_dirs,
    poc_test_emails,
    submission_email,
)
from src.constants import EmailConst, EmailReceiptStatus
from src.logging_utils import get_logger
from src.packaging import build_email_safe_zip

logger = get_logger()


def email_subject() -> str:
    """Canonical Outlook / MIME subject line for this submission."""
    return f"{EmailConst.SUBJECT_PREFIX}{candidate_name()}"

INLINE_PNGS = [
    ("Pivot_Monthly_Performance.png", "monthly_pivot"),
    ("Pivot_Country_Performance.png", "country_pivot"),
]


def _fmt_money(value: float | None) -> str:
    """Format GBP for email HTML (or n/a)."""
    if value is None:
        return "n/a"
    return f"£{value:,.2f}"


def _resolve_workbook_mb(workbook_mb: float | None) -> float | None:
    """Use caller-supplied size, else stat the default output workbook if present."""
    if workbook_mb is not None:
        return workbook_mb
    wb = OUTPUT_DIR / OUTPUT_XLSX_NAME
    return round(wb.stat().st_size / 1e6, 1) if wb.exists() else None


def _email_exec_summary_items(summary: dict) -> list[str]:
    """``<li>`` rows for the executive-summary section of the HTML body."""
    inv = summary.get("distinct_invoices")
    cust = summary.get("distinct_customers")
    inv_s = f"{inv:,}" if isinstance(inv, int) else str(inv)
    cust_s = f"{cust:,}" if isinstance(cust, int) else str(cust)
    items = [
        f"<li><b>Total revenue (Net)</b>: {_fmt_money(summary.get('email_total_revenue'))} "
        f"<i>(Gross − confirmed-return value only)</i></li>",
    ]
    net_after = summary.get("net_revenue_after_cancellations")
    if net_after is not None:
        items.append(
            f"<li><b>Net after cancellations</b>: {_fmt_money(net_after)} "
            f"<i>(also subtracts cancellation value £"
            f"{summary.get('cancellation_value', 0):,.2f})</i></li>"
        )
    ret_lines = summary.get("return_line_rows")
    conf = summary.get("confirmed_return_rows")
    adj_rows = summary.get("stock_adjustment_rows")
    adj_units = summary.get("adjustment_units")
    if ret_lines is not None:
        items.append(
            f"<li><b>Return lines</b>: {ret_lines:,} "
            f"<i>(is_return_line = non-C &amp; Qty&lt;0; "
            f"confirmed={conf:,}, stock adjustments={adj_rows:,}; "
            f"Top-10 ranked by return_line_rate)</i></li>"
        )
    if adj_units:
        items.append(
            f"<li><b>Stock adjustment units</b>: {adj_units:,.0f} "
            f"<i>(return-line subtype with Price≤0; not deducted from Net)</i></li>"
        )
    items.extend(
        [
            f"<li><b>Distinct invoices</b>: {inv_s}</li>",
            f"<li><b>Distinct customers</b>: {cust_s}</li>",
            f"<li><b>Top product</b>: {summary.get('top_product_stockcode')} — "
            f"{summary.get('top_product_description')} "
            f"({_fmt_money(summary.get('top_product_revenue'))})</li>",
            f"<li><b>Top non-UK country</b>: {summary.get('top_non_uk_country')} "
            f"({_fmt_money(summary.get('top_non_uk_country_revenue'))})</li>",
            f"<li><b>Best complete sales month</b>: {summary.get('best_complete_month')}</li>",
        ]
    )
    return items


def _email_inline_png_parts(*, embed_images: bool, cid_images: bool) -> list[str]:
    """HTML fragments for the two required inline Pivot PNGs."""
    parts: list[str] = []
    img_style = (
        f"max-width:{EmailConst.INLINE_IMG_MAX_WIDTH_PX}px;"
        f"border:1px solid {EmailConst.IMG_BORDER_COLOR};"
    )
    for filename, cid in INLINE_PNGS:
        path = PNG_DIR / filename
        if not path.exists():
            raise FileNotFoundError(path)
        parts.append(f"<p><b>{filename}</b></p>")
        if cid_images:
            parts.append(
                f"<img src='cid:{cid}' alt='{cid}' "
                f"style='max-width:{EmailConst.INLINE_IMG_MAX_WIDTH_PX}px;'/>"
            )
        elif embed_images:
            b64 = base64.b64encode(path.read_bytes()).decode("ascii")
            parts.append(
                f"<img src='data:image/png;base64,{b64}' alt='{cid}' "
                f"style='{img_style}'/>"
            )
        else:
            # Browser dry-run: relative file next to preview HTML.
            parts.append(f"<img src='{filename}' alt='{cid}' style='{img_style}'/>")
    return parts


def build_email_html(
    summary: dict,
    *,
    embed_images: bool = True,
    workbook_mb: float | None = None,
    cid_images: bool = False,
) -> str:
    """
    HTML body with exactly two Pivot screenshots inline.

    embed_images=True  → data-URI (Outlook send path; images visible in Sent Items)
    embed_images=False → relative filenames (browser dry-run preview next to copied PNGs)
    """
    xlsx_mb = _resolve_workbook_mb(workbook_mb)

    parts: list[str] = [
        "<html><body style='font-family:Aptos,Arial,sans-serif;font-size:12pt;'>",
        "<p>Dear Hiring Team,</p>",
        "<p>Please find my Python take-home assessment submission. "
        "All analysis, Excel PivotTables, screenshots, and this email were "
        "produced by Python automation.</p>",
    ]
    if xlsx_mb is not None and xlsx_mb > ATTACHMENT_LIMIT_MB:
        parts.append(
            "<p><b>Attachment note:</b> <code>Retail_Analysis.xlsx</code> is approximately "
            f"<b>{xlsx_mb:.0f}&nbsp;MB</b> (≈1M-row Clean_Data + native Pivot caches) and "
            "exceeds typical mailbox limits, so it is <b>not attached to this email</b>. "
            "The complete workbook is in the full submission ZIP. "
            "An alternate delivery channel requires confirmation.</p>"
        )
    parts.extend(["<h3>Executive summary</h3>", "<ul>"])
    parts.extend(_email_exec_summary_items(summary))
    parts.extend(
        [
            "</ul>",
            "<h3>Selected PivotTable screenshots</h3>",
            "<p>Monthly Performance and Country Performance.</p>",
        ]
    )
    parts.extend(_email_inline_png_parts(embed_images=embed_images, cid_images=cid_images))
    parts.extend(
        [
            f"<p>Kind regards,<br/>{candidate_name()}</p>",
            "</body></html>",
        ]
    )
    return "\n".join(parts)


def _write_cid_draft(path: Path, summary: dict, workbook: Path) -> None:
    """Write a local MIME draft for reviewing inline image structure."""
    message = EmailMessage()
    message["Subject"] = email_subject()
    message["X-Unsent"] = "1"
    message.set_content("Open this draft as HTML to view the executive summary and PivotTables.")
    size = workbook.stat().st_size
    message.add_alternative(
        build_email_html(summary, workbook_mb=size / 1e6, cid_images=True), subtype="html"
    )
    html_part = message.get_payload()[1]
    for filename, cid in INLINE_PNGS:
        html_part.add_related(
            (PNG_DIR / filename).read_bytes(), maintype="image", subtype="png",
            cid=f"<{cid}>", filename=filename, disposition="inline",
        )
    if size <= ATTACHMENT_LIMIT_BYTES:
        message.add_attachment(
            workbook.read_bytes(), maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="Retail_Analysis.xlsx",
        )
    path.write_bytes(message.as_bytes())


def write_dry_run_preview(summary: dict, workbook: Path, zip_path: Path | None) -> Path:
    """
    Write browser preview HTML, send-style HTML, CID .eml draft, and meta JSON
    under ``output/email_preview/``. Does not send mail. Returns preview HTML path.
    """
    ensure_dirs()
    preview_dir = OUTPUT_DIR / EMAIL_PREVIEW_DIRNAME
    preview_dir.mkdir(parents=True, exist_ok=True)

    wb_mb = round(workbook.stat().st_size / 1e6, 1) if workbook.exists() else None
    # Preview: relative <img src> + copied PNGs (opens in browser). Send: data-URI embeds.
    html_send = build_email_html(summary, embed_images=True, workbook_mb=wb_mb)
    html_preview = build_email_html(summary, embed_images=False, workbook_mb=wb_mb)
    for filename, _cid in INLINE_PNGS:
        src = PNG_DIR / filename
        shutil.copy2(src, preview_dir / filename)

    html_path = preview_dir / "submission_email_preview.html"
    html_path.write_text(html_preview, encoding="utf-8")
    (preview_dir / "submission_email_send.html").write_text(html_send, encoding="utf-8")

    if workbook.is_file():
        _write_cid_draft(preview_dir / "submission_email_cid.eml", summary, workbook)

    meta = {
        "subject": email_subject(),
        "from_account": "active Outlook profile (no passwords in source; uses signed-in Outlook)",
        "to_options": {
            "test": poc_test_emails(),
            "submission": submission_email(),
        },
        "body_metrics": {
            "total_revenue_net": summary.get("email_total_revenue"),
            "net_after_cancellations": summary.get("net_revenue_after_cancellations"),
            "distinct_invoices": summary.get("distinct_invoices"),
            "distinct_customers": summary.get("distinct_customers"),
            "top_product": summary.get("top_product_stockcode"),
            "top_non_uk_country": summary.get("top_non_uk_country"),
            "best_complete_month": summary.get("best_complete_month"),
        },
        "body_strategy": (
            "Inline: HTML data-URI for exactly two Pivot PNGs (Outlook). "
            "Also attach those PNGs as files (Gmail may strip data-URIs). "
            "The CLI blocks oversized workbook sends. CID .eml is available for offline review."
        ),
        "inline_images": [p for p, _ in INLINE_PNGS],
        "preview_images": "relative filenames beside submission_email_preview.html",
        "attachments": {
            "workbook_mb": wb_mb,
            "workbook_attached_when_mb_le": ATTACHMENT_LIMIT_MB,
            "zip": None
            if zip_path is None or not zip_path.exists()
            else {
                "path": str(zip_path),
                "mb": round(zip_path.stat().st_size / 1e6, 2),
            },
        },
        "send_html_bytes": len(html_send.encode("utf-8")),
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    (preview_dir / "submission_email_meta.json").write_text(
        json.dumps(meta, indent=2), encoding="utf-8"
    )
    logger.info("Email dry-run preview written | path=%s", html_path)
    return html_path


def _collect_attachments(
    workbook: Path,
    zip_path: Path | None,
    *,
    attach_workbook: bool,
    attach_zip: bool,
    attach_pngs: bool,
) -> tuple[list[Path], list[str]]:
    """
    Choose Outlook attachments under ATTACHMENT_LIMIT_BYTES.

    Returns (paths, human notes). May build an email-safe ZIP when the full ZIP is too large.
    """
    attachments: list[Path] = []
    notes: list[str] = []

    if attach_pngs:
        for name, _ in INLINE_PNGS:
            path = PNG_DIR / name
            if not path.exists():
                raise FileNotFoundError(path)
            attachments.append(path)

    if attach_workbook and workbook.exists():
        if workbook.stat().st_size <= ATTACHMENT_LIMIT_BYTES:
            attachments.append(workbook)
        else:
            notes.append(
                f"Full workbook not emailed ({workbook.stat().st_size / 1e6:.0f} MB); "
                "see complete ZIP / portal package."
            )

    if attach_zip and zip_path and zip_path.exists():
        if zip_path.stat().st_size <= ATTACHMENT_LIMIT_BYTES:
            attachments.append(zip_path)
        else:
            safe = build_email_safe_zip(
                PNG_DIR,
                AUDIT_DIR,
                summary_json=OUTPUT_DIR / ANALYSIS_SUMMARY_JSON,
            )
            if safe.stat().st_size <= ATTACHMENT_LIMIT_BYTES:
                attachments.append(safe)
                notes.append(f"Attached email-safe ZIP ({safe.name}).")

    return attachments, notes


def _write_receipt(path: Path, receipt: dict) -> None:
    path.write_text(json.dumps(receipt, indent=2), encoding="utf-8")


def _outlook_send_messages(
    *,
    subject: str,
    html: str,
    recipients: list[str],
    attachments: list[Path],
    sent_to: list[str],
) -> None:
    """Send one Outlook message per recipient; append each success onto ``sent_to``."""
    outlook = app("Microsoft Outlook")
    for to in recipients:
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={k.subject: subject, k.content: html},
        )
        msg.make(
            new=k.to_recipient,
            with_properties={k.email_address: {k.address: to}},
        )
        for path in attachments:
            msg.make(new=k.attachment, with_properties={k.file: Alias(str(path))})
        msg.send()
        sent_to.append(to)


def send_submission_email(
    summary: dict,
    workbook: Path,
    recipients: list[str],
    zip_path: Path | None = None,
    *,
    send: bool = False,
    attach_workbook: bool = True,
    attach_zip: bool = True,
    attach_pngs: bool = True,
) -> dict:
    """
    Build preview artifacts; optionally send via macOS Outlook (appscript).

    Default ``send=False`` is dry-run only. Returns a receipt dict with status
    (dry_run / sent_api_returned / failed). Never retries automatically.
    """
    ensure_dirs()
    if not recipients:
        raise ValueError("No recipients provided")
    cleaned: list[str] = []
    for addr in recipients:
        text = str(addr).strip()
        if not text or "@" not in text or text.startswith("@") or text.endswith("@"):
            raise ValueError(f"Invalid email recipient: {addr!r}")
        cleaned.append(text)
    recipients = cleaned

    subject = email_subject()
    wb_mb = round(workbook.stat().st_size / 1e6, 1) if workbook.exists() else None
    html = build_email_html(summary, embed_images=True, workbook_mb=wb_mb)
    preview = write_dry_run_preview(summary, workbook, zip_path)

    receipt = {
        "status": EmailReceiptStatus.DRY_RUN.value,
        "submission_id": f"mail-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
        "subject": subject,
        "recipients": recipients,
        "preview_html": str(preview),
        "workbook_mb": round(workbook.stat().st_size / 1e6, 2),
        "zip_mb": None
        if zip_path is None or not zip_path.exists()
        else round(zip_path.stat().st_size / 1e6, 2),
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    receipt_path = OUTPUT_DIR / EMAIL_PREVIEW_DIRNAME / "submission_receipt.json"

    if not send:
        _write_receipt(receipt_path, receipt)
        logger.info("Email dry-run only | send=False preview=%s", preview)
        return receipt

    attachments, notes = _collect_attachments(
        workbook,
        zip_path,
        attach_workbook=attach_workbook,
        attach_zip=attach_zip,
        attach_pngs=attach_pngs,
    )
    if notes:
        html = html.replace(
            "</body></html>",
            "<p><i>" + " ".join(notes) + "</i></p></body></html>",
        )

    receipt["status"] = EmailReceiptStatus.PENDING.value
    receipt["attachments"] = [str(p) for p in attachments]
    receipt["notes"] = notes
    _write_receipt(receipt_path, receipt)

    sent_to: list[str] = []
    try:
        _outlook_send_messages(
            subject=subject,
            html=html,
            recipients=recipients,
            attachments=attachments,
            sent_to=sent_to,
        )
    except Exception as exc:  # noqa: BLE001 — Outlook/appscript COM surface is broad
        # Convert send failure into a receipt (caller maps status→exit code); do not raise.
        receipt["status"] = EmailReceiptStatus.FAILED.value
        receipt["error"] = f"{type(exc).__name__}: {exc}"
        receipt["sent_to"] = sent_to
        receipt["failed_at"] = datetime.now().isoformat(timespec="seconds")
        _write_receipt(receipt_path, receipt)
        logger.exception("Outlook send failed | sent_to=%s", sent_to)
        return receipt

    receipt["status"] = EmailReceiptStatus.SENT.value
    receipt["sent_at"] = datetime.now().isoformat(timespec="seconds")
    receipt["sent_to"] = sent_to
    receipt["warning"] = (
        "Do not auto-retry. Check Outlook Sent Items (inline images) and recipient inbox. "
        "Gmail may strip data-URI images — PNG file attachments are the fallback."
    )
    _write_receipt(receipt_path, receipt)
    logger.info("Outlook send returned | sent_to=%s", sent_to)
    return receipt
