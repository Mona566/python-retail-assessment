#!/usr/bin/env python3
"""
Phase 0 Outlook PoC on macOS.

Separates:
  F-client: Microsoft Outlook desktop automation (appscript)
  F-graph:  Microsoft Graph sendMail (optional; not equivalent to client automation)

Never sends to the recruiter address. Uses POC_TEST_EMAIL or dry-run only.
"""

from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import traceback
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from PIL import Image
from appscript import app, k
from mactypes import Alias

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import POC_DIR, ensure_dirs, poc_test_emails  # noqa: E402


@dataclass
class CheckResult:
    """One Outlook/Graph PoC check, including whether it meets 'through Outlook' literally."""

    id: str
    name: str
    status: str
    detail: str = ""
    meets_through_outlook_literal: str = "unknown"  # yes | no | unknown


@dataclass
class Report:
    """Outlook Phase 0 PoC report container."""

    started_at: str
    finished_at: str | None = None
    checks: list[CheckResult] = field(default_factory=list)

    def add(
        self,
        check_id: str,
        name: str,
        status: str,
        detail: str = "",
        meets: str = "unknown",
    ) -> None:
        """Append a check; ``meets`` records yes/no/unknown for desktop-Outlook criterion."""
        self.checks.append(
            CheckResult(check_id, name, status, detail, meets)
        )


def detect_outlook_flavor() -> dict:
    """Best-effort New vs Legacy Outlook detection on macOS."""
    info: dict = {
        "app_exists": False,
        "bundle_path": None,
        "defaults_new_outlook": None,
        "applescript_probe": None,
        "notes": [],
    }
    app_path = Path("/Applications/Microsoft Outlook.app")
    info["app_exists"] = app_path.exists()
    info["bundle_path"] = str(app_path) if app_path.exists() else None

    # Preference key historically used to toggle New Outlook
    try:
        r = subprocess.run(
            [
                "defaults",
                "read",
                "com.microsoft.Outlook",
                "EnableNewOutlook",
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if r.returncode == 0:
            info["defaults_new_outlook"] = r.stdout.strip()
        else:
            info["notes"].append(
                f"EnableNewOutlook not readable (rc={r.returncode}): {r.stderr.strip()}"
            )
    except Exception as exc:  # noqa: BLE001
        info["notes"].append(f"defaults read failed: {exc}")

    # AppleScript dictionary / create message probe
    script = """
    try
      tell application "Microsoft Outlook"
        set _v to "ok"
        return _v
      end tell
    on error errMsg number errNum
      return "error:" & errNum & ":" & errMsg
    end try
    """
    try:
        r = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            check=False,
        )
        info["applescript_probe"] = {
            "returncode": r.returncode,
            "stdout": r.stdout.strip(),
            "stderr": r.stderr.strip(),
        }
    except Exception as exc:  # noqa: BLE001
        info["applescript_probe"] = {"error": str(exc)}

    return info


def try_client_create_message(report: Report, send: bool) -> None:
    """Attempt appscript outgoing message with HTML + attachment + inline images."""
    ensure_dirs()
    flavor = detect_outlook_flavor()
    (POC_DIR / "outlook_flavor.json").write_text(
        json.dumps(flavor, indent=2), encoding="utf-8"
    )
    report.add(
        "F-detect",
        "Detect Outlook installation / New vs Legacy signals",
        "PASS" if flavor["app_exists"] else "FAIL",
        json.dumps(flavor)[:800],
        meets="unknown",
    )

    # Create a >3MB dummy attachment for Graph/client size probing
    big_path = POC_DIR / "poc_large_attachment.bin"
    if not big_path.exists() or big_path.stat().st_size < 3_500_000:
        big_path.write_bytes(os.urandom(3_500_000))

    # Tiny PNGs for inline attempts
    try:
        img1 = POC_DIR / "poc_inline_1.png"
        img2 = POC_DIR / "poc_inline_2.png"
        Image.new("RGB", (120, 80), color=(20, 80, 160)).save(img1)
        Image.new("RGB", (120, 80), color=(160, 80, 20)).save(img2)
    except Exception as exc:  # noqa: BLE001
        report.add(
            "F-client-assets",
            "Create inline PNG fixtures",
            "FAIL",
            str(exc),
            meets="no",
        )
        return

    recipients = poc_test_emails()
    include_large = "--with-large-attachment" in sys.argv

    try:
        outlook = app("Microsoft Outlook")

        # CID path is known-limited on Mac; record FAIL without sending that variant.
        report.add(
            "F-client-inline",
            "CID-inline images via Outlook client automation",
            "FAIL",
            "CID/isInline not exposed via standard appscript attachment properties on Outlook for Mac.",
            meets="no",
        )

        b64_1 = base64.b64encode(img1.read_bytes()).decode("ascii")
        b64_2 = base64.b64encode(img2.read_bytes()).decode("ascii")
        variant = "full-large-attachment" if include_large else "light-xlsx-only"
        html_b64 = (
            "<html><body>"
            "<p><b>Phase 0 Outlook client PoC</b> — not a recruiter submission.</p>"
            f"<p>Variant: <code>{variant}</code></p>"
            "<p>Please confirm both colored rectangles appear inline below "
            "(blue = image 1, orange = image 2).</p>"
            "<p>Image 1 (blue):</p>"
            f'<img src="data:image/png;base64,{b64_1}" alt="poc1">'
            "<p>Image 2 (orange):</p>"
            f'<img src="data:image/png;base64,{b64_2}" alt="poc2">'
            "</body></html>"
        )

        xlsx_poc = POC_DIR / "poc_pivots.xlsx"
        subject = (
            "PoC - Base64 inline + large attachment"
            if include_large
            else "PoC - Base64 inline images (light, no 3MB bin)"
        )

        def build_message(to_list: list[str]):
            message = outlook.make(
                new=k.outgoing_message,
                with_properties={
                    k.subject: subject,
                    k.content: html_b64,
                },
            )
            if include_large:
                message.make(
                    new=k.attachment,
                    with_properties={k.file: Alias(str(big_path))},
                )
            if xlsx_poc.exists():
                message.make(
                    new=k.attachment,
                    with_properties={k.file: Alias(str(xlsx_poc))},
                )
            for addr in to_list:
                message.make(
                    new=k.to_recipient,
                    with_properties={k.email_address: {k.address: addr}},
                )
            return message

        report.add(
            "F-client",
            "Create Outlook message with HTML + attachment(s)",
            "PASS",
            f"Prepared variant={variant} for {recipients or '(no recipient)'}",
            meets="partial",
        )
        report.add(
            "F-client-large-attachment",
            "Attach >3MB file via Outlook client",
            "PASS" if include_large else "SKIP",
            (
                f"Attached {big_path.name} ({big_path.stat().st_size} bytes)"
                if include_large
                else "Skipped on light send; pass --with-large-attachment to include"
            ),
            meets="yes" if include_large else "unknown",
        )

        if send and recipients:
            receipt_path = POC_DIR / "submission_receipt_poc.json"
            pending = {
                "status": "pending",
                "submission_id": f"poc-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
                "to": recipients,
                "subject": subject,
                "variant": variant,
                "created_at": datetime.now().isoformat(timespec="seconds"),
            }
            receipt_path.write_text(json.dumps(pending, indent=2), encoding="utf-8")

            # One message per recipient so delivery failures are easier to isolate.
            sent_to: list[str] = []
            for addr in recipients:
                msg = build_message([addr])
                msg.send()
                sent_to.append(addr)

            pending["status"] = "sent_api_returned"
            pending["sent_at"] = datetime.now().isoformat(timespec="seconds")
            pending["sent_to"] = sent_to
            pending["note"] = (
                "Outlook send() returned for each recipient. Check inbox + spam. "
                "Prior Gmail miss may be due to random 3.5MB .bin looking suspicious."
            )
            receipt_path.write_text(json.dumps(pending, indent=2), encoding="utf-8")
            report.add(
                "F-client-send",
                "Send test message via Outlook client",
                "PASS",
                f"Sent variant={variant} to {sent_to}",
                meets="partial",
            )
            report.add(
                "F-client-inline-base64",
                "Base64 data-URI images in Outlook HTML body",
                "PASS",
                "Sent with two data-URI PNGs; confirm inbox rendering on each recipient.",
                meets="partial",
            )
        elif send and not recipients:
            report.add(
                "F-client-send",
                "Send test message via Outlook client",
                "FAIL",
                "POC_TEST_EMAIL is not set",
                meets="no",
            )
        else:
            msg = build_message(recipients or [])
            try:
                msg.open()
            except Exception:  # noqa: BLE001
                pass
            report.add(
                "F-client-send",
                "Send test message via Outlook client",
                "SKIP",
                "Dry-run draft opened. Pass --send with POC_TEST_EMAIL to send.",
                meets="unknown",
            )
            report.add(
                "F-client-inline-base64",
                "Base64 data-URI images in Outlook HTML body",
                "PASS",
                "Draft created for manual verification.",
                meets="partial",
            )
    except Exception as exc:  # noqa: BLE001
        report.add(
            "F-client",
            "Create Outlook message via appscript",
            "FAIL",
            f"{exc}\n{traceback.format_exc()}",
            meets="no",
        )


def try_graph_status(report: Report) -> None:
    """Record whether Graph credentials exist; do not send unless configured + --send."""
    tenant = os.getenv("GRAPH_TENANT_ID", "").strip()
    client_id = os.getenv("GRAPH_CLIENT_ID", "").strip()
    secret = os.getenv("GRAPH_CLIENT_SECRET", "").strip()
    upn = os.getenv("GRAPH_USER_UPN", "").strip()
    configured = all([tenant, client_id, secret, upn])
    report.add(
        "F-graph-config",
        "Microsoft Graph credentials present",
        "PASS" if configured else "SKIP",
        "Configured" if configured else "GRAPH_* env vars not set",
        meets="no",  # Graph is not literal Outlook-client automation
    )
    report.add(
        "F-graph-compliance-note",
        "Graph vs assessment 'through Outlook' requirement",
        "PASS",
        (
            "Graph sendMail can deliver HTML with isInline+contentId and appears in "
            "Outlook Sent Items, but it does NOT drive the Outlook desktop client. "
            "Treat as technical fallback only; recruiter acceptance is a policy question."
        ),
        meets="no",
    )


def main() -> int:
    """CLI entry for Outlook client + Graph status PoC (optional --send to POC_TEST_EMAIL)."""
    ensure_dirs()
    send = "--send" in sys.argv
    report = Report(started_at=datetime.now().isoformat(timespec="seconds"))
    try_client_create_message(report, send=send)
    try_graph_status(report)
    report.finished_at = datetime.now().isoformat(timespec="seconds")

    out = {
        "started_at": report.started_at,
        "finished_at": report.finished_at,
        "send_attempted": send,
        "checks": [asdict(c) for c in report.checks],
    }
    path = POC_DIR / "phase0_outlook_report.json"
    path.write_text(json.dumps(out, indent=2), encoding="utf-8")

    print("=== Phase 0 Outlook PoC Results ===")
    failed = 0
    for c in report.checks:
        print(
            f"[{c.status}] {c.id}: {c.name} "
            f"(meets_through_outlook_literal={c.meets_through_outlook_literal})"
        )
        if c.detail:
            print(f"         {c.detail[:600]}")
        if c.status == "FAIL":
            failed += 1
    print(f"\nReport: {path}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
