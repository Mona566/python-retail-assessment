"""Build the submission ZIP package."""

from __future__ import annotations

import json
import zipfile
from datetime import datetime
from pathlib import Path

from src.config import OUTPUT_DIR, PROJECT_ROOT, candidate_name, ensure_dirs  # noqa: F401
from src.constants import Artifacts, PackageKind, PackagingConst
from src.data_processing import sha256_file
from src.logging_utils import get_logger

logger = get_logger()


def _add_common_files(
    zf: zipfile.ZipFile,
    manifest: dict,
    png_dir: Path,
    audit_dir: Path,
    summary_json: Path | None,
) -> None:
    """Add source, tests, PNGs, audit files, and summary into the ZIP + manifest."""
    for rel in [
        "main.py",
        "path_setup.py",
        "pytest.ini",
        "requirements.txt",
        "README.md",
        ".env.example",
        "docs/PIPELINE_FLOW.md",
        "docs/ASSIGNMENT_WALKTHROUGH.md",
    ]:
        path = PROJECT_ROOT / rel
        if path.exists():
            zf.write(path, arcname=rel)
            manifest["files"][rel] = {
                "sha256": sha256_file(path),
                "bytes": path.stat().st_size,
            }

    for path in sorted([*(PROJECT_ROOT / "src").rglob("*.py"), *(PROJECT_ROOT / "tests").glob("test_*.py"), PROJECT_ROOT / "tests/conftest.py"]):
        if not path.is_file() or "__pycache__" in path.parts:
            continue
        rel = path.relative_to(PROJECT_ROOT).as_posix()
        zf.write(path, arcname=rel)
        manifest["files"][rel] = {
            "sha256": sha256_file(path),
            "bytes": path.stat().st_size,
        }

    for png in sorted(png_dir.glob(PackagingConst.PIVOT_PNG_GLOB)):
        arc = f"{PackagingConst.PNG_ARC_PREFIX}{png.name}"
        zf.write(png, arcname=arc)
        manifest["files"][arc] = {
            "sha256": sha256_file(png),
            "bytes": png.stat().st_size,
        }

    for path in sorted(audit_dir.glob("*")):
        if path.is_file() and path.suffix in {".json", ".csv"}:
            arc = f"{PackagingConst.AUDIT_ARC_PREFIX}{path.name}"
            zf.write(path, arcname=arc)
            manifest["files"][arc] = {
                "sha256": sha256_file(path),
                "bytes": path.stat().st_size,
            }

    if summary_json and summary_json.exists():
        zf.write(summary_json, arcname=summary_json.name)
        manifest["files"][summary_json.name] = {
            "sha256": sha256_file(summary_json),
            "bytes": summary_json.stat().st_size,
        }


def build_submission_zip(
    workbook: Path,
    png_dir: Path,
    audit_dir: Path,
    summary_json: Path | None = None,
) -> Path:
    """
    Full submission package including Retail_Analysis.xlsx (may be large).

    Bundles source, tests, PNGs, audit JSON/CSV, summary, and the workbook.
    """
    ensure_dirs()
    name = candidate_name().replace(" ", "")
    zip_path = OUTPUT_DIR / f"{name}{PackagingConst.SUBMISSION_ZIP_SUFFIX}"
    if zip_path.exists():
        zip_path.unlink()

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "candidate_name": candidate_name(),
        "package": PackageKind.FULL.value,
        "files": {},
    }

    if not workbook.exists():
        raise FileNotFoundError(workbook)

    with zipfile.ZipFile(
        zip_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=PackagingConst.ZIP_COMPRESS_LEVEL,
    ) as zf:
        _add_common_files(zf, manifest, png_dir, audit_dir, summary_json)
        zf.write(workbook, arcname=workbook.name)
        manifest["files"][workbook.name] = {
            "sha256": sha256_file(workbook),
            "bytes": workbook.stat().st_size,
        }
        zf.writestr(Artifacts.MANIFEST_JSON, json.dumps(manifest, indent=2))

    logger.info(
        "Submission ZIP written | path=%s mb=%.1f",
        zip_path,
        zip_path.stat().st_size / 1e6,
    )
    return zip_path


def build_email_safe_zip(
    png_dir: Path,
    audit_dir: Path,
    summary_json: Path | None = None,
) -> Path:
    """
    Mail-safe package: source + PNGs + audit/summary, WITHOUT the ~100MB+ workbook.
    Use when Outlook/mailbox limits block attaching Retail_Analysis.xlsx.
    """
    ensure_dirs()
    name = candidate_name().replace(" ", "")
    zip_path = OUTPUT_DIR / f"{name}{PackagingConst.EMAIL_SAFE_ZIP_SUFFIX}"
    if zip_path.exists():
        zip_path.unlink()

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "candidate_name": candidate_name(),
        "package": PackageKind.EMAIL_SAFE.value,
        "note": (
            "Full Retail_Analysis.xlsx omitted due to mailbox size limits. "
            "Included: source, README, Pivot PNGs, audit/summary."
        ),
        "files": {},
    }

    with zipfile.ZipFile(
        zip_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=PackagingConst.ZIP_COMPRESS_LEVEL,
    ) as zf:
        _add_common_files(zf, manifest, png_dir, audit_dir, summary_json)
        zf.writestr(Artifacts.MANIFEST_JSON, json.dumps(manifest, indent=2))
        zf.writestr(
            Artifacts.ATTACHMENT_NOTE_TXT,
            (
                "Retail_Analysis.xlsx was omitted from this email package because "
                "the full workbook exceeds typical mailbox attachment limits.\n"
                "The complete submission ZIP (with workbook) is available locally as "
                f"{name}{PackagingConst.SUBMISSION_ZIP_SUFFIX}.\n"
            ),
        )

    logger.info(
        "Email-safe ZIP written | path=%s mb=%.1f",
        zip_path,
        zip_path.stat().st_size / 1e6,
    )
    return zip_path
