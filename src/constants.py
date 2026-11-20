"""
Named constants and closed enums for policy strings, artifacts, and tuning knobs.

Finite label sets use ``str, Enum`` so they compare equal to their string values
(``PipelineState.DONE == "DONE"``) while remaining type-safe and discoverable.
Note: on Python 3.11+, ``str(SomeEnum.MEMBER)`` may render as ``"SomeEnum.MEMBER"``;
pass ``.value`` into Excel/pandas sheet names and other APIs that stringify.
Open-ended filenames / numeric knobs stay as namespace classes.
"""

from __future__ import annotations

from enum import Enum
from typing import Final


class PipelineState(str, Enum):
    """Explicit pipeline state-machine labels (main + PipelineRun)."""

    INIT = "INIT"
    READY = "READY"
    PREPARE = "PREPARE"
    ANALYSE = "ANALYSE"
    EXCEL = "EXCEL"
    VALIDATE = "VALIDATE"
    PREVIEW_LOAD = "PREVIEW_LOAD"
    PACKAGE = "PACKAGE"
    EMAIL_PREVIEW = "EMAIL_PREVIEW"
    EMAIL_SEND = "EMAIL_SEND"
    DONE = "DONE"
    FAILED = "FAILED"
    ABORTED = "ABORTED"

    @classmethod
    def terminal(cls) -> frozenset[PipelineState]:
        """States that end a run (success, failure, or abort)."""
        return frozenset({cls.DONE, cls.FAILED, cls.ABORTED})


# Alias kept for ``state not in PipelineState.TERMINAL`` call sites.
PipelineState.TERMINAL = PipelineState.terminal()  # type: ignore[attr-defined]


class StepStatus(str, Enum):
    """Per-step / whole-run status values written to the run audit JSON."""

    STARTED = "started"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    SKIPPED = "skipped"
    RUNNING = "running"  # whole-run status while in progress


class ReturnLineClass(str, Enum):
    """Subtype label stored in ``return_line_class`` for non-cancellation negatives."""

    CONFIRMED = "confirmed_return"
    ADJUSTMENT = "stock_adjustment"


class PackageKind(str, Enum):
    """Submission ZIP package flavour in the ZIP manifest."""

    FULL = "full"
    EMAIL_SAFE = "email_safe_no_workbook"


class EmailReceiptStatus(str, Enum):
    """Outlook send receipt ``status`` field."""

    DRY_RUN = "dry_run"
    PENDING = "pending"
    SENT = "sent_api_returned"
    FAILED = "failed"


class AnomalyBucket(str, Enum):
    """Quota buckets used when selecting Top-N anomalies."""

    MONETARY = "monetary"
    ADJUSTMENT = "adjustment"
    CUSTOMER = "customer"


class AnomalyReason(str, Enum):
    """Stable reason tokens stored in the anomalies ``reason`` column (``;``-joined)."""

    CUSTOMER_FREQUENCY = "abnormal_customer_invoice_frequency"
    STOCK_ADJUSTMENT_QTY = "extreme_stock_adjustment_qty"
    LARGE_QUANTITY = "unusually_large_quantity"
    HIGH_PRICE = "unusually_high_price"
    NEGATIVE_PRICE = "negative_price"
    CONFIRMED_RETURN = "extreme_confirmed_return"
    CANCELLATION = "extreme_cancellation"
    LARGE_LINE_VALUE = "unusually_large_line_value"
    ZERO_QUANTITY = "zero_quantity"


class Artifacts:
    """Stable filenames under data/ / output/ / audit/."""

    CLEAN_PARQUET: Final[str] = "clean_transactions.parquet"
    PREP_SUMMARY_JSON: Final[str] = "prep_summary.json"
    ANALYSIS_SUMMARY_JSON: Final[str] = "analysis_summary.json"
    OFFICE_BUILD_JSON: Final[str] = "office_build.json"
    DATA_QUALITY_JSON: Final[str] = "data_quality_summary.json"
    REVENUE_RECON_JSON: Final[str] = "revenue_reconciliation.json"
    CHANGED_RECORDS_CSV: Final[str] = "changed_records.csv"
    REMOVED_CLONES_CSV: Final[str] = "removed_cross_sheet_clones.csv"
    ANOMALY_THRESHOLDS_JSON: Final[str] = "anomaly_thresholds.json"
    MANIFEST_JSON: Final[str] = "manifest.json"
    ATTACHMENT_NOTE_TXT: Final[str] = "ATTACHMENT_NOTE.txt"
    PIPELINE_LOG: Final[str] = "pipeline.log"
    PIPELINE_RUN_LATEST_JSON: Final[str] = "pipeline_run_latest.json"


class Sheets(str, Enum):
    """Workbook sheet names used outside PIVOT_SPECS."""

    CLEAN_DATA = "Clean_Data"
    SUMMARY = "Summary"
    ANOMALIES = "Anomalies"
    RETURNS_SOURCE = "Returns_Source"


class BusinessRules:
    """Domain policy literals shared by prep / analysis / validation."""

    HOME_COUNTRY: Final[str] = "United Kingdom"
    CANCELLATION_INVOICE_PREFIX: Final[str] = "C"
    FLOAT_EQ_TOLERANCE: Final[float] = 1e-6
    PIVOT_GRAND_TOTAL_LABEL: Final[str] = "grand total"
    RETURN_CLASS_CONFIRMED: Final[str] = ReturnLineClass.CONFIRMED.value
    RETURN_CLASS_ADJUSTMENT: Final[str] = ReturnLineClass.ADJUSTMENT.value


class PackagingConst:
    """ZIP naming and compression."""

    SUBMISSION_ZIP_SUFFIX: Final[str] = "_PythonAssessment.zip"
    EMAIL_SAFE_ZIP_SUFFIX: Final[str] = "_PythonAssessment_EmailSafe.zip"
    ZIP_COMPRESS_LEVEL: Final[int] = 9
    PIVOT_PNG_GLOB: Final[str] = "Pivot_*.png"
    PNG_ARC_PREFIX: Final[str] = "png/"
    AUDIT_ARC_PREFIX: Final[str] = "audit/"


class EmailConst:
    """Outlook / preview email presentation."""

    SUBJECT_PREFIX: Final[str] = "Python Assessment Submission - "
    PREVIEW_DIRNAME: Final[str] = "email_preview"
    INLINE_IMG_MAX_WIDTH_PX: Final[int] = 900
    IMG_BORDER_COLOR: Final[str] = "#ccc"


class HashingConst:
    """File hashing I/O."""

    SHA256_CHUNK_BYTES: Final[int] = 1024 * 1024


class ExcelPngConst:
    """Pivot screenshot export tuning."""

    MIN_BYTES: Final[int] = 5000
    WORKBOOK_SETTLE_SEC: Final[float] = 1.5
    SHEET_SETTLE_SEC: Final[float] = 0.4
    FORCE_FILL_RGB: Final[tuple[int, int, int]] = (255, 255, 255)
    FORCE_FONT_RGB: Final[tuple[int, int, int]] = (0, 0, 0)


class PivotLayoutConst:
    """Native PivotTable column-width heuristics for PNG readability."""

    COL_WIDTH_VALUE: Final[int] = 15
    COL_WIDTH_ROW_LABEL: Final[int] = 22
    COL_WIDTH_MIN: Final[int] = 18
    HEADER_WIDTH_PAD: Final[int] = 8
    PRODUCT_WIDTH_PAD: Final[int] = 9
    HEADER_ROW_INDEX: Final[int] = 3  # A3 destination → style row
