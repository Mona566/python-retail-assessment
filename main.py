#!/usr/bin/env python3
"""
Retail Transaction Analysis pipeline.

Default: prepare data, analyse, build workbook, export PNGs, build ZIP,
and write an Outlook email dry-run preview (does NOT send).

Send only with --send after reviewing output/email_preview/.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from path_setup import ensure_repo_root_on_sys_path

ensure_repo_root_on_sys_path()

from src.analysis import AnalysisBundle, run_analyses  # noqa: E402
from src.config import (  # noqa: E402
    ANALYSIS_SUMMARY_JSON,
    ATTACHMENT_LIMIT_BYTES,
    AUDIT_DIR,
    OFFICE_BUILD_JSON,
    OUTPUT_DIR,
    OUTPUT_XLSX_NAME,
    PNG_DIR,
    candidate_name,
    ensure_dirs,
    poc_test_emails,
)
from src.constants import EmailReceiptStatus, PipelineState  # noqa: E402
from src.data_processing import (  # noqa: E402
    StaleCleanCacheError,
    find_input_workbook,
    load_prepared,
    prepare_dataset,
    sha256_file,
    validate_enrich_schema,
)
from src.email_automation import (  # noqa: E402
    email_subject,
    send_submission_email,
    write_dry_run_preview,
)
from src.excel_automation import build_workbook  # noqa: E402
from src.logging_utils import PipelineRun  # noqa: E402
from src.packaging import build_email_safe_zip, build_submission_zip  # noqa: E402
from src.validation import REQUIRED_PNGS, run_validations, validate_pngs  # noqa: E402


@dataclass
class PipelinePaths:
    """Resolved artifact locations for one pipeline run."""

    workbook: Path
    summary_path: Path
    build_manifest: Path
    artifact_paths: list[Path]


@dataclass
class StageResult:
    """Outcome of a pipeline stage; early_exit stops main after finish()."""

    exit_code: int = 0
    early_exit: bool = False
    summary: dict[str, Any] | None = None
    workbook: Path | None = None
    zip_path: Path | None = None
    email_zip: Path | None = None
    preview: Path | None = None


def parse_args() -> argparse.Namespace:
    """
    Parse CLI flags for prep/excel/zip/email modes.

    Notable: --reuse-clean, --skip-excel/--skip-png/--skip-zip,
    --email-preview-only, --send / --send-to.
    """
    p = argparse.ArgumentParser(description="Online Retail II analysis pipeline")
    p.add_argument("--input", type=Path, default=None)
    p.add_argument("--output", type=Path, default=None)
    p.add_argument(
        "--reuse-clean",
        action="store_true",
        help="Reuse data/processed/clean_transactions.parquet",
    )
    p.add_argument("--skip-excel", action="store_true")
    p.add_argument("--skip-png", action="store_true")
    p.add_argument("--skip-zip", action="store_true")
    p.add_argument(
        "--email-preview-only",
        action="store_true",
        help="Assume workbook/PNGs already exist; only refresh email preview + ZIP metadata",
    )
    p.add_argument(
        "--send",
        action="store_true",
        help=(
            "Actually send via Outlook (use sparingly). "
            "Requires --send-to, or falls back to first POC_TEST_EMAIL only "
            "(never auto-uses SUBMISSION_EMAIL)."
        ),
    )
    p.add_argument(
        "--send-to",
        type=str,
        default=None,
        help="Recipient for --send (single address). Prefer a self-test inbox.",
    )
    return p.parse_args()


def _pipeline_paths(args: argparse.Namespace) -> PipelinePaths:
    """Build workbook/summary/manifest paths from CLI args and config defaults."""
    workbook = Path(args.output) if args.output else OUTPUT_DIR / OUTPUT_XLSX_NAME
    summary_path = OUTPUT_DIR / ANALYSIS_SUMMARY_JSON
    return PipelinePaths(
        workbook=workbook,
        summary_path=summary_path,
        build_manifest=OUTPUT_DIR / OFFICE_BUILD_JSON,
        artifact_paths=[workbook, summary_path, *(PNG_DIR / name for name in REQUIRED_PNGS)],
    )


def _run_prepare(
    run: PipelineRun,
    args: argparse.Namespace,
) -> tuple[StageResult, Any, dict]:
    """Load or prepare clean data. Returns (stage, clean, audit) on success."""
    run.transition(PipelineState.READY, PipelineState.PREPARE, reason="data preparation")
    with run.step(
        "prepare_data",
        reuse_clean=args.reuse_clean,
        input=str(args.input) if args.input else None,
    ) as prep_meta:
        try:
            prep = load_prepared() if args.reuse_clean else prepare_dataset(args.input)
        except StaleCleanCacheError as exc:
            run.logger.error("%s", exc)
            run.transition(PipelineState.PREPARE, PipelineState.FAILED, reason="cache_schema_mismatch")
            return StageResult(exit_code=1, early_exit=True), None, {}
        if args.reuse_clean and prep.input_sha256 != sha256_file(
            find_input_workbook(args.input)
        ):
            run.logger.error(
                "Cached clean data does not match the input workbook. "
                "Run without --reuse-clean."
            )
            run.transition(PipelineState.PREPARE, PipelineState.FAILED, reason="cache_sha_mismatch")
            return StageResult(exit_code=1, early_exit=True), None, {}
        if args.reuse_clean:
            # Defense in depth if load_prepared is stubbed in tests.
            schema_errors = validate_enrich_schema(prep.clean)
            if schema_errors:
                run.logger.error("%s", schema_errors[0])
                run.transition(PipelineState.PREPARE, PipelineState.FAILED, reason="cache_schema_mismatch")
                return StageResult(exit_code=1, early_exit=True), None, {}
        prep_meta.update(
            {
                "rows": len(prep.clean),
                "reuse_clean": args.reuse_clean,
                "input_sha256": prep.input_sha256[:12] + "…",
            }
        )
    return StageResult(), prep.clean, prep.audit_summary


def _run_analyse(
    run: PipelineRun,
    clean,
    summary_path: Path,
) -> tuple[StageResult, AnalysisBundle]:
    """Run analyses and write analysis_summary.json."""
    run.transition(PipelineState.PREPARE, PipelineState.ANALYSE, reason="business analyses")
    with run.step("analyse", rows=len(clean)) as analyse_meta:
        bundle = run_analyses(clean)
        summary_path.write_text(
            json.dumps(bundle.summary, indent=2, default=str), encoding="utf-8"
        )
        analyse_meta.update(
            {
                "net_revenue": round(float(bundle.summary["net_revenue"]), 2),
                "distinct_invoices": bundle.summary["distinct_invoices"],
                "top_product": bundle.summary["top_product_stockcode"],
                "top_country": bundle.summary["top_non_uk_country"],
                "best_month": bundle.summary["best_complete_month"],
                "summary_path": str(summary_path),
            }
        )
        run.logger.info(
            "Net £%s | invoices=%s | top product=%s | top country=%s | best month=%s",
            f"{bundle.summary['net_revenue']:,.2f}",
            bundle.summary["distinct_invoices"],
            bundle.summary["top_product_stockcode"],
            bundle.summary["top_non_uk_country"],
            bundle.summary["best_complete_month"],
        )
    return StageResult(summary=bundle.summary), bundle


def _run_excel(
    run: PipelineRun,
    args: argparse.Namespace,
    clean,
    bundle: AnalysisBundle,
    audit: dict,
    workbook: Path,
) -> tuple[StageResult, Path]:
    """Build workbook + optional PNGs, or skip when --skip-excel."""
    if args.skip_excel:
        run.skip("build_workbook", reason="--skip-excel")
        run.transition(PipelineState.ANALYSE, PipelineState.VALIDATE, reason="excel skipped")
        return StageResult(workbook=workbook), workbook

    run.transition(PipelineState.ANALYSE, PipelineState.EXCEL, reason="workbook + pivots + png")
    with run.step(
        "build_workbook",
        export_png=not args.skip_png,
        output=str(workbook),
    ) as excel_meta:
        workbook = build_workbook(
            clean,
            bundle,
            audit,
            output_path=workbook,
            export_png=not args.skip_png,
        )
        excel_meta.update(
            {
                "workbook": str(workbook),
                "workbook_mb": round(workbook.stat().st_size / 1e6, 2),
                "export_png": not args.skip_png,
            }
        )
    return StageResult(workbook=workbook), workbook


def _run_validate(
    run: PipelineRun,
    args: argparse.Namespace,
    clean,
    audit: dict,
    bundle: AnalysisBundle,
    workbook: Path,
    paths: PipelinePaths,
) -> StageResult:
    """Validate outputs; write office_build.json when a full package is expected."""
    if not args.skip_excel:
        run.transition(PipelineState.EXCEL, PipelineState.VALIDATE, reason="post-excel checks")

    with run.step(
        "validate",
        check_workbook=not args.skip_excel,
        check_pngs=(not args.skip_excel and not args.skip_png),
    ) as val_meta:
        errors = run_validations(
            clean,
            audit,
            bundle,
            workbook if not args.skip_excel else None,
            check_pngs=(not args.skip_excel and not args.skip_png),
        )
        val_meta["error_count"] = len(errors)
        if errors:
            run.logger.error("VALIDATION FAIL | count=%s", len(errors))
            run.transition(PipelineState.VALIDATE, PipelineState.FAILED, reason="validation_errors")
            return StageResult(exit_code=1, early_exit=True)

    if args.skip_excel or args.skip_png:
        run.logger.info(
            "Analysis finished. Run with Excel and PNG export enabled "
            "to build the submission."
        )
        run.transition(PipelineState.VALIDATE, PipelineState.DONE, reason="partial_run_no_package")
        return StageResult(
            exit_code=0,
            early_exit=True,
            summary=bundle.summary,
            workbook=workbook,
        )

    paths.build_manifest.write_text(
        json.dumps([sha256_file(path) for path in paths.artifact_paths], indent=2),
        encoding="utf-8",
    )
    return StageResult(summary=bundle.summary, workbook=workbook)


def _run_full_pipeline(
    run: PipelineRun,
    args: argparse.Namespace,
    paths: PipelinePaths,
) -> StageResult:
    """Prepare → analyse → excel → validate (stops early on failure or partial skip)."""
    stage, clean, audit = _run_prepare(run, args)
    if stage.early_exit:
        return stage

    stage, bundle = _run_analyse(run, clean, paths.summary_path)
    stage, workbook = _run_excel(run, args, clean, bundle, audit, paths.workbook)
    return _run_validate(run, args, clean, audit, bundle, workbook, paths)


def _run_preview_only(
    run: PipelineRun,
    paths: PipelinePaths,
) -> StageResult:
    """Load existing workbook/summary/PNGs for --email-preview-only (hash-checked)."""
    run.transition(PipelineState.READY, PipelineState.PREVIEW_LOAD, reason="email-preview-only")
    with run.step("load_existing_artifacts") as load_meta:
        if not paths.summary_path.exists() or not paths.workbook.exists():
            run.logger.error(
                "email-preview-only requires existing workbook + analysis_summary.json"
            )
            run.transition(PipelineState.PREVIEW_LOAD, PipelineState.FAILED, reason="missing_artifacts")
            return StageResult(exit_code=1, early_exit=True)
        errors = validate_pngs()
        if errors:
            run.logger.error("Invalid PNGs: %s", errors)
            run.transition(PipelineState.PREVIEW_LOAD, PipelineState.FAILED, reason="invalid_pngs")
            return StageResult(exit_code=1, early_exit=True)
        if not paths.build_manifest.exists() or json.loads(
            paths.build_manifest.read_text(encoding="utf-8")
        ) != [sha256_file(path) for path in paths.artifact_paths]:
            run.logger.error(
                "Workbook, summary and screenshots must match one validated build. "
                "Rebuild first."
            )
            run.transition(PipelineState.PREVIEW_LOAD, PipelineState.FAILED, reason="build_manifest_mismatch")
            return StageResult(exit_code=1, early_exit=True)
        summary = json.loads(paths.summary_path.read_text(encoding="utf-8"))
        load_meta.update(
            {
                "workbook": str(paths.workbook),
                "summary_path": str(paths.summary_path),
            }
        )
        run.logger.info("Loaded existing summary/workbook | mode=email_preview_only")
    return StageResult(summary=summary, workbook=paths.workbook)


def _package_and_preview(
    run: PipelineRun,
    args: argparse.Namespace,
    paths: PipelinePaths,
    summary: dict[str, Any],
    workbook: Path,
) -> StageResult:
    """Build submission ZIP (and email-safe ZIP if needed) plus dry-run email preview."""
    zip_path = None
    email_zip = None

    if not args.skip_zip and workbook.exists():
        run.transition(run.state, PipelineState.PACKAGE, reason="submission zip")
        with run.step("package_zip", workbook=str(workbook)) as zip_meta:
            zip_path = build_submission_zip(
                workbook=workbook,
                png_dir=PNG_DIR,
                audit_dir=AUDIT_DIR,
                summary_json=paths.summary_path if paths.summary_path.exists() else None,
            )
            if (
                zip_path.stat().st_size > ATTACHMENT_LIMIT_BYTES
                or workbook.stat().st_size > ATTACHMENT_LIMIT_BYTES
            ):
                email_zip = build_email_safe_zip(
                    PNG_DIR,
                    AUDIT_DIR,
                    summary_json=paths.summary_path if paths.summary_path.exists() else None,
                )
            else:
                email_zip = zip_path
            zip_meta.update(
                {
                    "full_zip_mb": round(zip_path.stat().st_size / 1e6, 2),
                    "email_safe": email_zip != zip_path,
                    "email_zip_mb": round(email_zip.stat().st_size / 1e6, 2),
                }
            )
    else:
        run.skip(
            "package_zip",
            reason="--skip-zip" if args.skip_zip else "workbook missing",
        )

    run.transition(run.state, PipelineState.EMAIL_PREVIEW, reason="dry-run preview")
    with run.step("email_preview") as mail_meta:
        preview = write_dry_run_preview(summary, workbook, email_zip or zip_path)
        mail_meta["preview_html"] = str(preview)
        run.logger.info("Email preview ready | path=%s", preview)
        run.logger.info(
            "Email subject | subject=%s",
            email_subject(),
        )
        if zip_path:
            run.logger.info(
                "Full ZIP ready | path=%s mb=%.1f",
                zip_path,
                zip_path.stat().st_size / 1e6,
            )
        if email_zip and email_zip != zip_path:
            run.logger.info(
                "Email-safe ZIP ready | path=%s mb=%.1f",
                email_zip,
                email_zip.stat().st_size / 1e6,
            )

    return StageResult(
        summary=summary,
        workbook=workbook,
        zip_path=zip_path,
        email_zip=email_zip,
        preview=preview,
    )


def _resolve_send_recipients(
    run: PipelineRun,
    args: argparse.Namespace,
) -> tuple[list[str] | None, StageResult | None]:
    """Prefer --send-to; else first POC_TEST_EMAIL. Never auto-use SUBMISSION_EMAIL."""

    def _valid_addr(raw: str) -> str | None:
        text = raw.strip()
        if not text or "@" not in text or text.startswith("@") or text.endswith("@"):
            return None
        return text

    if args.send_to:
        addr = _valid_addr(args.send_to)
        if addr is None:
            run.logger.error("Invalid --send-to address | value=%r", args.send_to)
            run.transition(
                PipelineState.EMAIL_SEND, PipelineState.FAILED, reason="invalid_recipient"
            )
            return None, StageResult(exit_code=2, early_exit=True)
        return [addr], None
    if poc_test_emails():
        addr = _valid_addr(poc_test_emails()[0])
        if addr is None:
            run.logger.error("Invalid POC_TEST_EMAIL | value=%r", poc_test_emails()[0])
            run.transition(
                PipelineState.EMAIL_SEND, PipelineState.FAILED, reason="invalid_recipient"
            )
            return None, StageResult(exit_code=2, early_exit=True)
        recipients = [addr]
        run.logger.warning(
            "No --send-to; using POC_TEST_EMAIL | recipient=%s", recipients[0]
        )
        return recipients, None
    run.logger.error(
        "No recipient | hint=pass --send-to or set POC_TEST_EMAIL "
        "(SUBMISSION_EMAIL is never an automatic send target)"
    )
    run.transition(PipelineState.EMAIL_SEND, PipelineState.FAILED, reason="no_recipient")
    return None, StageResult(exit_code=2, early_exit=True)


def _maybe_send(
    run: PipelineRun,
    args: argparse.Namespace,
    summary: dict[str, Any],
    workbook: Path,
    preview: Path | None,
    email_zip: Path | None = None,
) -> StageResult:
    """Finish dry-run, or send one Outlook message when --send is set.

    Uses the email-safe (or full) ZIP from packaging when available so the
    attachment strategy matches the preview path.
    """
    if not args.send:
        run.logger.info(
            "Pipeline OK (no email sent). Review %s then run with --send --send-to <addr>",
            preview,
        )
        run.transition(run.state, PipelineState.DONE, reason="dry_run_complete")
        return StageResult(exit_code=0, preview=preview)

    run.transition(PipelineState.EMAIL_PREVIEW, PipelineState.EMAIL_SEND, reason="--send")
    workbook_over_limit = workbook.stat().st_size > ATTACHMENT_LIMIT_BYTES
    attachable_zip = (
        email_zip
        if email_zip is not None
        and email_zip.exists()
        and email_zip.stat().st_size <= ATTACHMENT_LIMIT_BYTES
        else None
    )
    # Oversized workbook is only a hard failure when there is no email ZIP to attach.
    if workbook_over_limit and attachable_zip is None:
        run.logger.error(
            "The required workbook exceeds the attachment limit and no email-safe "
            "ZIP is available. Confirm an alternate delivery channel."
        )
        run.transition(PipelineState.EMAIL_SEND, PipelineState.FAILED, reason="workbook_too_large")
        return StageResult(exit_code=2, early_exit=True)
    if workbook_over_limit:
        run.logger.warning(
            "Workbook exceeds attachment limit (%.1f MB); attaching email ZIP only: %s",
            workbook.stat().st_size / 1e6,
            attachable_zip,
        )

    recipients, err = _resolve_send_recipients(run, args)
    if err is not None:
        return err

    run.logger.warning(
        "Sending ONE Outlook message to %s — avoid repeats to reduce block risk",
        recipients,
    )
    with run.step(
        "email_send",
        recipients=recipients,
        zip_path=str(attachable_zip) if attachable_zip else None,
        attach_zip=attachable_zip is not None,
    ) as send_meta:
        receipt = send_submission_email(
            summary,
            workbook,
            recipients=recipients,
            zip_path=attachable_zip,
            send=True,
            attach_workbook=not workbook_over_limit,
            attach_zip=attachable_zip is not None,
        )
        send_meta["receipt_status"] = receipt.get("status")
        run.logger.info("Send receipt | status=%s", receipt.get("status"))
        if receipt.get("status") == EmailReceiptStatus.FAILED:
            run.transition(PipelineState.EMAIL_SEND, PipelineState.FAILED, reason="outlook_send_failed")
            return StageResult(exit_code=3, early_exit=True)

    run.transition(PipelineState.EMAIL_SEND, PipelineState.DONE, reason="sent")
    return StageResult(exit_code=0)


def main() -> int:
    """
    CLI entry: orchestrate pipeline stages with PipelineRun audit/timing.

    Returns a process exit code (0 success; non-zero on validation/send failures).
    """
    args = parse_args()
    ensure_dirs()

    mode = "email_preview_only" if args.email_preview_only else "full_pipeline"
    run = PipelineRun(
        mode=mode,
        cli_flags={
            "reuse_clean": args.reuse_clean,
            "skip_excel": args.skip_excel,
            "skip_png": args.skip_png,
            "skip_zip": args.skip_zip,
            "email_preview_only": args.email_preview_only,
            "send": args.send,
            "has_send_to": bool(args.send_to),
            "has_input": args.input is not None,
            "has_output": args.output is not None,
        },
    )
    paths = _pipeline_paths(args)
    exit_code = 0

    try:
        if args.email_preview_only:
            stage = _run_preview_only(run, paths)
        else:
            stage = _run_full_pipeline(run, args, paths)
        if stage.early_exit:
            return stage.exit_code

        summary = stage.summary
        workbook = stage.workbook or paths.workbook
        assert summary is not None

        packaged = _package_and_preview(run, args, paths, summary, workbook)
        sent = _maybe_send(
            run,
            args,
            summary,
            workbook,
            packaged.preview,
            email_zip=packaged.email_zip,
        )
        exit_code = sent.exit_code
        return exit_code

    except Exception:
        # Top-level safety net: record FAILED + traceback, then non-zero exit.
        # Do not catch BaseException (KeyboardInterrupt / SystemExit).
        exit_code = exit_code or 1
        if run.state not in PipelineState.TERMINAL:
            run.transition("*", PipelineState.FAILED, reason="unhandled_exception")
        run.logger.exception("Pipeline aborted with unhandled exception")
        return exit_code
    finally:
        run.finish(exit_code)


if __name__ == "__main__":
    raise SystemExit(main())
