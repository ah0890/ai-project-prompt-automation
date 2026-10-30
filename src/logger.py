"""CSV processing log and application logging, with secret redaction.

The CSV log is append-only: each row is written, flushed, and fsynced
immediately, so an interruption can lose at most the row being written and
never corrupts earlier rows.
"""

from __future__ import annotations

import csv
import logging
import os
import re
from pathlib import Path

from src.progress_manager import now_iso

CSV_COLUMNS = [
    "Project Number",
    "Excel Row Number",
    "Project Name",
    "Project Identifier",
    "Status",
    "Timestamp",
    "Error Message",
]

_SECRET_PATTERNS = [
    re.compile(r"sk-ant-[A-Za-z0-9_\-]{8,}"),
    re.compile(r"(?i)(x-api-key|authorization|api[_-]?key)(\s*[:=]\s*)([^\s,;\"']+)"),
]


def redact(text: str, secrets: tuple[str, ...] = ()) -> str:
    text = str(text)
    for secret in secrets:
        if secret and len(secret) >= 8:
            text = text.replace(secret, "[REDACTED]")
    text = _SECRET_PATTERNS[0].sub("[REDACTED]", text)
    return _SECRET_PATTERNS[1].sub(lambda m: f"{m.group(1)}{m.group(2)}[REDACTED]", text)


def _csv_safe(value: object) -> str:
    """Neutralize spreadsheet formula injection when the log is opened in Excel."""
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ")
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


class ProcessingLog:
    def __init__(self, path: Path, secrets: tuple[str, ...] = ()):
        self.path = path
        self.secrets = secrets
        self._ensure_header()

    def _ensure_header(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists() and self.path.stat().st_size > 0:
            with self.path.open("r", encoding="utf-8-sig", newline="") as handle:
                first = next(csv.reader(handle), None)
            if first == CSV_COLUMNS:
                return
            # Unknown layout: keep the old file rather than mixing formats.
            legacy = self.path.with_name(f"{self.path.stem}.previous-{os.getpid()}{self.path.suffix}")
            os.replace(self.path, legacy)
        with self.path.open("w", encoding="utf-8-sig", newline="") as handle:
            csv.writer(handle).writerow(CSV_COLUMNS)
            handle.flush()
            os.fsync(handle.fileno())

    def write(
        self,
        status: str,
        project_number: int | str = "",
        excel_row: int | str = "",
        project_name: str = "",
        project_id: str = "",
        error: str = "",
    ) -> None:
        row = [
            project_number if project_number == "" else f"{int(project_number):03d}",
            excel_row,
            project_name,
            project_id,
            status,
            now_iso(),
            redact(error, self.secrets),
        ]
        with self.path.open("a", encoding="utf-8", newline="") as handle:
            csv.writer(handle).writerow([_csv_safe(v) for v in row])
            handle.flush()
            os.fsync(handle.fileno())


class RedactingFilter(logging.Filter):
    def __init__(self, secrets: tuple[str, ...] = ()):
        super().__init__()
        self.secrets = secrets

    def filter(self, record: logging.LogRecord) -> bool:
        record.msg = redact(record.getMessage(), self.secrets)
        record.args = ()
        return True


def setup_logging(log_file: Path, secrets: tuple[str, ...] = ()) -> logging.Logger:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("prompt_generator")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()
    handler = logging.FileHandler(log_file, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    handler.addFilter(RedactingFilter(secrets))
    logger.addHandler(handler)
    return logger
