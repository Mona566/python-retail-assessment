"""Tests for pipeline run tracking, step timing, and audit JSON."""

from __future__ import annotations

import json
import logging
import logging.handlers
import time

import pytest

from src.logging_utils import PipelineRun, setup_logging, timed_phase


@pytest.fixture(autouse=True)
def isolate_logs(monkeypatch, tmp_path):
    from src import config
    from src import logging_utils

    monkeypatch.setattr(config, "LOG_DIR", tmp_path)
    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    # Reset logger handlers/filters so setup_logging re-binds to tmp_path.
    logger = logging.getLogger("retail_assessment")
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()
    for filt in list(logger.filters):
        logger.removeFilter(filt)


def test_pipeline_run_records_steps_and_timing(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)

    run = PipelineRun(mode="test", cli_flags={"dry": True}, run_id="testrun01")
    run.transition("READY", "WORK", reason="begin")
    with run.step("alpha", n=1) as meta:
        meta["rows"] = 3
    run.skip("beta", reason="not needed")
    with run.step("gamma"):
        pass
    run.transition("WORK", "DONE", reason="complete")
    audit_path = run.finish(0)

    assert audit_path.exists()
    payload = json.loads(audit_path.read_text(encoding="utf-8"))
    assert payload["run_id"] == "testrun01"
    assert payload["status"] == "succeeded"
    assert payload["exit_code"] == 0
    assert payload["state"] == "DONE"
    assert payload["total_duration_sec"] >= 0
    names = [s["name"] for s in payload["steps"]]
    assert names == ["alpha", "beta", "gamma"]
    assert payload["steps"][0]["status"] == "succeeded"
    assert payload["steps"][0]["details"]["rows"] == 3
    assert payload["steps"][1]["status"] == "skipped"
    assert "alpha" in payload["timing_summary"]
    assert "beta" not in payload["timing_summary"]
    latest = tmp_path / "pipeline_run_latest.json"
    assert latest.exists()


def test_pipeline_run_step_failure_is_recorded(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)

    run = PipelineRun(mode="test", run_id="failrun01")
    with pytest.raises(ValueError, match="boom"):
        with run.step("broken"):
            raise ValueError("boom")
    run.transition("*", "FAILED", reason="step_failed")
    audit_path = run.finish(1)

    payload = json.loads(audit_path.read_text(encoding="utf-8"))
    assert payload["status"] == "failed"
    assert payload["steps"][0]["status"] == "failed"
    assert "ValueError" in (payload["steps"][0]["error"] or "")


def test_timed_phase_logs_duration(caplog):
    setup_logging()
    with caplog.at_level(logging.INFO, logger="retail_assessment"):
        with timed_phase("unit_phase", x=1) as meta:
            meta["y"] = 2
    messages = " | ".join(r.message for r in caplog.records)
    assert "PHASE START unit_phase" in messages
    assert "PHASE OK unit_phase" in messages


def test_timed_phase_logs_rows_per_sec(caplog):
    setup_logging()
    with caplog.at_level(logging.INFO, logger="retail_assessment"):
        with timed_phase("throughput_phase", rows=10_000) as meta:
            time.sleep(0.01)
            meta["ok"] = 1
    ok_lines = [r.message for r in caplog.records if "PHASE OK throughput_phase" in r.message]
    assert ok_lines
    assert "rows_per_sec=" in ok_lines[0]


def test_log_format_includes_run_and_step(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    run = PipelineRun(mode="test", run_id="fmt01")
    memory = logging.handlers.MemoryHandler(capacity=100, flushLevel=logging.CRITICAL)
    memory.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] [run=%(run_id)s step=%(step)s] %(message)s")
    )
    run.logger.addHandler(memory)
    try:
        with run.step("fmt_step"):
            run.logger.info("hello-audit")
        formatted = [memory.format(r) for r in memory.buffer]
    finally:
        run.logger.removeHandler(memory)
        memory.close()
    assert any("[run=fmt01 step=fmt_step] hello-audit" in line for line in formatted)
    assert any("STEP OK fmt_step" in line for line in formatted)


def test_timed_phase_failure(caplog):
    setup_logging()
    with caplog.at_level(logging.ERROR, logger="retail_assessment"):
        with pytest.raises(RuntimeError, match="phase-boom"):
            with timed_phase("bad_phase"):
                raise RuntimeError("phase-boom")
    assert any("PHASE FAIL bad_phase" in r.message for r in caplog.records)


def test_transition_mismatch_and_finish_when_done(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    run = PipelineRun(mode="test", run_id="mismatch01")
    # READY already; claim we are WORK
    run.transition("WORK", "DONE", reason="forced")
    # finish when already DONE — should not re-transition via mismatch path for DONE
    path = run.finish(0)
    assert path.exists()
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["state"] == "DONE"


def test_finish_transitions_from_mid_state(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    run = PipelineRun(mode="test", run_id="mid01")
    run.transition("READY", "WORK", reason="go")
    path = run.finish(1)
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["state"] == "FAILED"
    assert payload["status"] == "failed"


def test_logging_conventions_documented_and_shared_logger():
    """Single logger name; structured prefixes; lazy %-formatting helpers."""
    import src.analysis as an
    import src.data_processing as dp
    import src.email_automation as ea
    import src.excel_automation as excel
    import src.packaging as pkg
    import src.validation as val
    from src.logging_utils import LOGGER_NAME, get_logger

    assert get_logger().name == LOGGER_NAME == "retail_assessment"
    for mod in (an, dp, ea, excel, pkg, val):
        assert mod.logger is get_logger() or mod.logger.name == LOGGER_NAME


def test_get_logger_and_json_safe_format_details():
    from pathlib import Path

    from src.logging_utils import _format_details, _json_safe, get_logger

    assert get_logger().name == "retail_assessment"
    assert _json_safe(None) is None
    assert _json_safe(True) is True
    assert _json_safe(Path("/tmp/x")) == "/tmp/x"
    assert _json_safe({"a": Path("b")}) == {"a": "b"}
    assert _json_safe([1, Path("p")]) == [1, "p"]
    assert _json_safe(object())  # falls back to str

    assert _format_details({"a": 1, "b": None, "c": "x"}) == "a=1 c=x"


def test_setup_logging_rebinds_file_handler(tmp_path, monkeypatch):
    from pathlib import Path

    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    logger = setup_logging(run_id="rebind1")
    first = tmp_path / "pipeline.log"
    assert first.exists() or any(isinstance(h, logging.FileHandler) for h in logger.handlers)

    other = tmp_path / "other_logs"
    other.mkdir()
    monkeypatch.setattr(logging_utils, "LOG_DIR", other)
    logger2 = setup_logging(run_id="rebind2")
    assert any(
        isinstance(h, logging.FileHandler)
        and Path(h.baseFilename).resolve() == (other / "pipeline.log").resolve()
        for h in logger2.handlers
    )


def test_step_failure_merges_extra_details(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    run = PipelineRun(mode="test", run_id="extradet")
    with pytest.raises(ValueError):
        with run.step("s") as meta:
            meta["partial"] = 1
            raise ValueError("x")
    assert run.steps[0].details.get("partial") == 1


def test_write_audit_default_total(tmp_path, monkeypatch):
    from src import logging_utils

    monkeypatch.setattr(logging_utils, "LOG_DIR", tmp_path)
    run = PipelineRun(mode="test", run_id="aud01")
    path = run.write_audit()
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert "total_duration_sec" in payload
