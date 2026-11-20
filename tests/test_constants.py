"""Sanity checks for grouped named constants and closed enums."""

from __future__ import annotations

from src.constants import (
    AnomalyBucket,
    AnomalyReason,
    Artifacts,
    BusinessRules,
    EmailConst,
    EmailReceiptStatus,
    ExcelPngConst,
    PackageKind,
    PackagingConst,
    PipelineState,
    ReturnLineClass,
    Sheets,
    StepStatus,
)
from src.metrics import CANCELLATION_INVOICE_PREFIX, HOME_COUNTRY


def test_constants_are_stable_and_reexported():
    assert HOME_COUNTRY == BusinessRules.HOME_COUNTRY == "United Kingdom"
    assert CANCELLATION_INVOICE_PREFIX == BusinessRules.CANCELLATION_INVOICE_PREFIX == "C"
    assert PackagingConst.SUBMISSION_ZIP_SUFFIX.endswith(".zip")
    assert EmailConst.SUBJECT_PREFIX.startswith("Python Assessment")
    assert ExcelPngConst.MIN_BYTES > 0
    assert Artifacts.ANALYSIS_SUMMARY_JSON.endswith(".json")


def test_closed_sets_are_str_enums():
    assert PipelineState.DONE == "DONE"
    assert PipelineState.DONE in PipelineState.TERMINAL
    assert "READY" not in PipelineState.TERMINAL

    assert StepStatus.SUCCEEDED == "succeeded"
    assert StepStatus.FAILED.value == "failed"

    assert ReturnLineClass.CONFIRMED == "confirmed_return"
    assert BusinessRules.RETURN_CLASS_CONFIRMED == ReturnLineClass.CONFIRMED.value

    assert PackageKind.FULL == "full"
    assert PackageKind.EMAIL_SAFE == "email_safe_no_workbook"

    assert EmailReceiptStatus.SENT == "sent_api_returned"
    assert EmailReceiptStatus.FAILED == "failed"

    assert AnomalyBucket.MONETARY == "monetary"
    assert set(AnomalyBucket) == {
        AnomalyBucket.MONETARY,
        AnomalyBucket.ADJUSTMENT,
        AnomalyBucket.CUSTOMER,
    }

    assert AnomalyReason.CUSTOMER_FREQUENCY == "abnormal_customer_invoice_frequency"
    assert AnomalyReason.ZERO_QUANTITY.value == "zero_quantity"
    assert Sheets.RETURNS_SOURCE.value == "Returns_Source"

    # str(Enum) is "Sheets.CLEAN_DATA" on 3.11+; Excel I/O must use .value
    assert Sheets.CLEAN_DATA.value == "Clean_Data"
    assert Sheets.CLEAN_DATA == "Clean_Data"  # equality still holds for str Enum
