"""
Shared logging setup, pipeline step tracking, and run-timing audit.

Logging conventions (keep call sites aligned):

- **Logger**: always ``get_logger()`` / ``LOGGER_NAME`` — never a second logger name.
- **Format**: ``%(asctime)s [%(levelname)s] [run=%(run_id)s step=%(step)s] %(message)s``
- **Lazy args**: ``logger.info("… %s", value)`` — do not use f-strings as the message.
- **Structured events** (grep-friendly prefixes):
  - ``STATE <from> → <to> | reason``
  - ``STEP START|OK|FAIL|SKIP <name> | key=value …``
  - ``PHASE START|OK|FAIL <name> | key=value …``
  - ``VALIDATION FAIL|OK`` for aggregate validation results
- **Levels**: DEBUG = cleanup/best-effort; INFO = progress; WARNING = recoverable
  policy risk; ERROR = user-visible failure; ``exception()`` only with a stack.
- **Details**: prefer ``key=value`` pairs (see ``_format_details``); separate clauses with `` | ``.
"""

from __future__ import annotations

import json
import logging
import resource
import sys
import time
import uuid
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Iterator

from src.config import LOG_DIR, ensure_dirs
from src.constants import Artifacts, PipelineState, StepStatus

LOGGER_NAME = "retail_assessment"

_run_id_var: ContextVar[str] = ContextVar("pipeline_run_id", default="-")
_step_var: ContextVar[str] = ContextVar("pipeline_step", default="-")


class _ContextFilter(logging.Filter):
    """Inject run_id / step from contextvars into every log record."""

    def filter(self, record: logging.LogRecord) -> bool:
        """Attach ``run_id`` and ``step`` attributes used by the log format string."""
        record.run_id = _run_id_var.get()  # type: ignore[attr-defined]
        record.step = _step_var.get()  # type: ignore[attr-defined]
        return True


@dataclass
class StepRecord:
    """One timed pipeline step for the JSON run audit."""

    name: str
    status: str
    started_at: str
    ended_at: str | None = None
    duration_sec: float | None = None
    details: dict[str, Any] = field(default_factory=dict)
    error: str | None = None


def _utc_now_iso() -> str:
    """UTC timestamp string for audit records."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def setup_logging(
    name: str = LOGGER_NAME,
    *,
    level: int = logging.INFO,
    run_id: str | None = None,
) -> logging.Logger:
    """
    Configure console + file handlers.

    Log line format includes run_id and current step for audit grepping:
      2026-... [INFO] [run=a1b2c3d4 step=prepare_data] message

    Rebinds the file handler when LOG_DIR changes (useful in tests).
    """
    ensure_dirs()
    logger = logging.getLogger(name)
    if run_id:
        _run_id_var.set(run_id)

    logger.setLevel(level)
    logger.propagate = False

    # Attach on the logger so every handler (incl. test caplog) sees run_id/step.
    if not any(isinstance(f, _ContextFilter) for f in logger.filters):
        logger.addFilter(_ContextFilter())

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] [run=%(run_id)s step=%(step)s] %(message)s"
    )
    log_file = (LOG_DIR / Artifacts.PIPELINE_LOG).resolve()

    has_stream = any(
        isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler)
        for h in logger.handlers
    )
    if not has_stream:
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)

    matched_file = False
    for handler in list(logger.handlers):
        if not isinstance(handler, logging.FileHandler):
            continue
        if Path(handler.baseFilename).resolve() == log_file:
            matched_file = True
            handler.setFormatter(fmt)
        else:
            logger.removeHandler(handler)
            handler.close()
    if not matched_file:
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

    return logger


def get_logger(name: str = LOGGER_NAME) -> logging.Logger:
    """Return the shared retail_assessment logger (call setup_logging first in CLI)."""
    return logging.getLogger(name)


class PipelineRun:
    """
    Tracks pipeline state transitions and per-step wall-clock timing.

    Writes a JSON audit to output/logs/pipeline_run_<run_id>.json on finish(),
    plus a stable symlink-style latest file pipeline_run_latest.json.
    """

    def __init__(
        self,
        *,
        mode: str,
        cli_flags: dict[str, Any] | None = None,
        run_id: str | None = None,
    ) -> None:
        """Start a run in INIT→READY; configures logging with this run_id."""
        self.run_id = run_id or uuid.uuid4().hex[:12]
        self.mode = mode
        self.cli_flags = cli_flags or {}
        self.started_at = _utc_now_iso()
        self.ended_at: str | None = None
        self.status = StepStatus.RUNNING
        self.exit_code: int | None = None
        self.state = PipelineState.INIT
        self.steps: list[StepRecord] = []
        self._t0 = time.perf_counter()
        self.logger = setup_logging(run_id=self.run_id)
        self.logger.info(
            "Pipeline run started | mode=%s | flags=%s",
            self.mode,
            self.cli_flags,
        )
        self.transition(
            PipelineState.INIT, PipelineState.READY, reason="run_created"
        )

    def transition(self, from_state: str, to_state: str, *, reason: str = "") -> None:
        """Record and log an explicit state-machine hop."""
        if self.state != from_state and from_state != "*":
            self.logger.warning(
                "State mismatch on transition: expected=%s actual=%s → %s (%s)",
                from_state,
                self.state,
                to_state,
                reason or "n/a",
            )
        prev = self.state
        self.state = to_state
        self.logger.info(
            "STATE %s → %s%s",
            prev,
            to_state,
            f" | {reason}" if reason else "",
        )

    @contextmanager
    def step(self, name: str, **details: Any) -> Iterator[dict[str, Any]]:
        """
        Time a named pipeline step. On success logs duration; on failure
        records the exception and re-raises.
        """
        token = _step_var.set(name)
        started = _utc_now_iso()
        t0 = time.perf_counter()
        record = StepRecord(
            name=name,
            status=StepStatus.STARTED,
            started_at=started,
            details={k: _json_safe(v) for k, v in details.items()},
        )
        self.steps.append(record)
        self.logger.info(
            "STEP START %s%s",
            name,
            f" | {_format_details(details)}" if details else "",
        )
        extra: dict[str, Any] = {}
        try:
            yield extra
        except Exception as exc:
            duration = time.perf_counter() - t0
            record.status = StepStatus.FAILED
            record.ended_at = _utc_now_iso()
            record.duration_sec = round(duration, 3)
            record.error = f"{type(exc).__name__}: {exc}"
            if extra:
                record.details.update({k: _json_safe(v) for k, v in extra.items()})
            self.logger.exception(
                "STEP FAIL %s | %.3fs | %s",
                name,
                duration,
                record.error,
            )
            raise
        else:
            duration = time.perf_counter() - t0
            record.status = StepStatus.SUCCEEDED
            record.ended_at = _utc_now_iso()
            record.duration_sec = round(duration, 3)
            if extra:
                record.details.update({k: _json_safe(v) for k, v in extra.items()})
            self.logger.info(
                "STEP OK %s | %.3fs%s",
                name,
                duration,
                f" | {_format_details(extra)}" if extra else "",
            )
        finally:
            _step_var.reset(token)

    def skip(self, name: str, *, reason: str, **details: Any) -> None:
        """Record a skipped step (duration 0) without failing the run."""
        now = _utc_now_iso()
        record = StepRecord(
            name=name,
            status=StepStatus.SKIPPED,
            started_at=now,
            ended_at=now,
            duration_sec=0.0,
            details={"reason": reason, **{k: _json_safe(v) for k, v in details.items()}},
        )
        self.steps.append(record)
        token = _step_var.set(name)
        try:
            self.logger.info("STEP SKIP %s | %s", name, reason)
        finally:
            _step_var.reset(token)

    def finish(self, exit_code: int, *, status: str | None = None) -> Path:
        """Close the run, log a timing summary, and persist JSON audit."""
        self.exit_code = exit_code
        self.status = status or (
            StepStatus.SUCCEEDED if exit_code == 0 else StepStatus.FAILED
        )
        self.ended_at = _utc_now_iso()
        total = round(time.perf_counter() - self._t0, 3)
        if self.state not in PipelineState.TERMINAL:
            self.transition(
                "*",
                PipelineState.DONE if exit_code == 0 else PipelineState.FAILED,
                reason="finish",
            )

        summary_lines = [
            f"{s.name}={s.status}:{s.duration_sec if s.duration_sec is not None else '-'}s"
            for s in self.steps
        ]
        rss_mb = _peak_rss_mb()
        self.logger.info(
            "Pipeline finished | status=%s exit=%s total=%.3fs%s | %s",
            self.status,
            exit_code,
            total,
            f" peak_rss_mb={rss_mb:.1f}" if rss_mb is not None else "",
            " | ".join(summary_lines) if summary_lines else "(no steps)",
        )
        return self.write_audit(total_sec=total, peak_rss_mb=rss_mb)

    def write_audit(
        self, *, total_sec: float | None = None, peak_rss_mb: float | None = None
    ) -> Path:
        """Persist pipeline_run_<id>.json and pipeline_run_latest.json under LOG_DIR."""
        ensure_dirs()
        if total_sec is None:
            total_sec = round(time.perf_counter() - self._t0, 3)
        payload = {
            "run_id": self.run_id,
            "mode": self.mode,
            "cli_flags": self.cli_flags,
            "status": _json_safe(self.status),
            "exit_code": self.exit_code,
            "state": _json_safe(self.state),
            "started_at": self.started_at,
            "ended_at": self.ended_at or _utc_now_iso(),
            "total_duration_sec": total_sec,
            "peak_rss_mb": peak_rss_mb,
            "steps": [_json_safe(asdict(s)) for s in self.steps],
            "timing_summary": {
                s.name: s.duration_sec
                for s in self.steps
                if s.duration_sec is not None and s.status != StepStatus.SKIPPED
            },
        }
        run_path = LOG_DIR / f"pipeline_run_{self.run_id}.json"
        latest_path = LOG_DIR / Artifacts.PIPELINE_RUN_LATEST_JSON
        text = json.dumps(payload, indent=2, ensure_ascii=False, default=str)
        run_path.write_text(text, encoding="utf-8")
        latest_path.write_text(text, encoding="utf-8")
        self.logger.info("Run audit written | path=%s", run_path)
        return run_path


def _peak_rss_mb() -> float | None:
    """Peak resident set size in MiB (best-effort; platform units differ)."""
    try:
        rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        # macOS: bytes; Linux: kilobytes
        if sys.platform == "darwin":
            return round(rss / (1024 * 1024), 1)
        return round(rss / 1024, 1)
    except Exception:  # noqa: BLE001 — metrics must never break the pipeline
        return None


@contextmanager
def timed_phase(
    name: str,
    *,
    logger: logging.Logger | None = None,
    **details: Any,
) -> Iterator[dict[str, Any]]:
    """
    Lightweight nested timer for sub-phases inside a module
    (e.g. load_raw / enrich / write_parquet) without registering a full PipelineRun step.

    Uses ``time.perf_counter()``. When ``rows`` / ``clean_rows`` is provided (start
    kwargs or yielded meta), logs ``rows_per_sec`` on success.
    """
    log = logger or get_logger()
    token = _step_var.set(name)
    t0 = time.perf_counter()
    log.info(
        "PHASE START %s%s",
        name,
        f" | {_format_details(details)}" if details else "",
    )
    extra: dict[str, Any] = {}
    try:
        yield extra
    except Exception as exc:
        duration = time.perf_counter() - t0
        log.exception(
            "PHASE FAIL %s | %.3fs | %s: %s",
            name,
            duration,
            type(exc).__name__,
            exc,
        )
        raise
    else:
        duration = time.perf_counter() - t0
        merged = {**details, **extra}
        rows = merged.get("rows")
        if rows is None:
            rows = merged.get("clean_rows")
        if isinstance(rows, (int, float)) and rows > 0 and duration > 0:
            merged.setdefault("rows_per_sec", round(float(rows) / duration, 1))
        # Always surface wall time; throughput appears when a row count was provided.
        log.info(
            "PHASE OK %s | %.3fs%s",
            name,
            duration,
            f" | {_format_details(merged)}" if merged else "",
        )
    finally:
        _step_var.reset(token)


def _json_safe(value: Any) -> Any:
    """Convert values to JSON-serializable forms for audit payloads and log details."""
    # Enum before str: ``str, Enum`` members are instances of str.
    if isinstance(value, Enum):
        return value.value
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(v) for v in value]
    return str(value)


def _format_details(details: dict[str, Any]) -> str:
    """Format ``key=value`` pairs for STEP/PHASE log lines (skips None)."""
    parts = []
    for key, value in details.items():
        if value is None:
            continue
        parts.append(f"{key}={value}")
    return " ".join(parts)
