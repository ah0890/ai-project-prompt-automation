"""Persistent progress, prompt cache, and project identity.

Two files under DATA_DIR:

- generated_prompts.json - every successfully generated prompt (the source of
  truth for the Word document; it is never discarded automatically).
- progress.json - per-project status, attempts, last error, and usage totals.
  It can be rebuilt from the prompt cache if it is ever corrupted.

Project identity:
- project_id is derived from the normalized Project Name + GitHub Repo Name,
  so inserting or reordering Excel rows does not trigger regeneration.
- content_hash covers all five source fields; when it differs from the cached
  hash the cached prompt is outdated and the project is regenerated.
- rules_sha256 records which prompt-rules.md produced each prompt. Prompts made
  under different rules are reported but only regenerated on request
  (--apply-new-rules), because regeneration costs money.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import tempfile
import time
from dataclasses import dataclass
from typing import Any
from pathlib import Path

from src.excel_reader import Project, SkippedRow

SCHEMA_VERSION = 1

STATUS_SUCCESS = "SUCCESS"
STATUS_FAILED = "FAILED"
STATUS_OUTDATED = "OUTDATED"


class ProgressError(Exception):
    """Persistent state cannot be used safely; the message is user-facing."""


@dataclass(frozen=True)
class TrackedProject:
    number: int  # 1-based position among eligible projects (PROJECT 001, ...)
    project: Project
    project_id: str
    content_hash: str

    @property
    def name(self) -> str:
        return self.project.project_name


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds")


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def content_hash(project: Project) -> str:
    payload = json.dumps({k: _norm(v) for k, v in project.fields().items()}, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def identity_key(project: Project) -> str:
    return f"{_norm(project.project_name).casefold()}|{_norm(project.github_repo_name).casefold()}"


def track_projects(projects: list[Project]) -> tuple[list[TrackedProject], list[SkippedRow]]:
    """Assign stable IDs; exact duplicate rows are skipped, not regenerated."""
    tracked: list[TrackedProject] = []
    skipped: list[SkippedRow] = []
    seen_hashes: dict[str, int] = {}
    identity_counts: dict[str, int] = {}
    for project in projects:
        digest = content_hash(project)
        if digest in seen_hashes:
            skipped.append(
                SkippedRow(
                    project.excel_row,
                    f"Duplicate of Excel row {seen_hashes[digest]} (identical project information)",
                    project.project_name,
                )
            )
            continue
        seen_hashes[digest] = project.excel_row
        key = identity_key(project)
        identity_counts[key] = identity_counts.get(key, 0) + 1
        base_id = "p-" + hashlib.sha256(key.encode("utf-8")).hexdigest()[:12]
        occurrence = identity_counts[key]
        project_id = base_id if occurrence == 1 else f"{base_id}-{occurrence}"
        tracked.append(TrackedProject(len(tracked) + 1, project, project_id, digest))
    return tracked, skipped


def atomic_write_text(path: Path, text: str, validate=None) -> None:
    """Write to a temp file in the same directory, validate, then replace."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        if validate:
            validate(tmp)
        replace_with_retry(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)


def replace_with_retry(src: Path, dst: Path, attempts: int = 5) -> None:
    """os.replace, retrying briefly for transient Windows locks (antivirus, indexer)."""
    for attempt in range(attempts):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == attempts - 1:
                raise
            time.sleep(0.2 * (attempt + 1))


def atomic_write_json(path: Path, data: dict) -> None:
    text = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=False) + "\n"

    def _validate(tmp: Path) -> None:
        json.loads(tmp.read_text(encoding="utf-8"))

    atomic_write_text(path, text, validate=_validate)


class ProgressManager:
    def __init__(self, progress_path: Path, prompts_path: Path):
        self.progress_path = progress_path
        self.prompts_path = prompts_path
        self.prompts: dict[str, dict] = {}
        self.progress: dict = {}
        self.warnings: list[str] = []

    # ------------------------------------------------------------------ load
    def load(self) -> None:
        self.prompts = self._load_prompts()
        self.progress = self._load_progress()

    def _load_prompts(self) -> dict[str, dict]:
        if not self.prompts_path.exists():
            return {}
        try:
            data = json.loads(self.prompts_path.read_text(encoding="utf-8"))
            prompts = data["prompts"]
            if not isinstance(prompts, dict):
                raise TypeError("'prompts' is not an object")
            return prompts
        except (OSError, ValueError, KeyError, TypeError) as exc:
            raise ProgressError(
                f"The prompt cache {self.prompts_path} is unreadable ({exc}).\n"
                "It holds previously generated (paid-for) prompts, so it was NOT overwritten.\n"
                "Restore it from a backup, or move it aside to start a fresh cache, then rerun."
            ) from exc

    def _load_progress(self) -> dict:
        fresh: dict[str, Any] = {"version": SCHEMA_VERSION, "projects": {}, "usage_totals": {}, "last_run": None}
        if not self.progress_path.exists():
            return fresh
        try:
            data = json.loads(self.progress_path.read_text(encoding="utf-8"))
            if not isinstance(data.get("projects"), dict):
                raise TypeError("'projects' is not an object")
            data.setdefault("usage_totals", {})
            data.setdefault("last_run", None)
            return data
        except (OSError, ValueError, TypeError, AttributeError) as exc:
            backup = self.progress_path.with_name(
                f"{self.progress_path.name}.corrupt-{dt.datetime.now():%Y%m%d-%H%M%S}"
            )
            os.replace(self.progress_path, backup)
            self.warnings.append(
                f"progress.json was unreadable ({exc}); moved to {backup.name} and rebuilt "
                "from the prompt cache. Cached prompts are unaffected."
            )
            for pid, entry in self.prompts.items():
                fresh["projects"][pid] = {
                    "status": STATUS_SUCCESS,
                    "excel_row": entry.get("excel_row"),
                    "project_name": entry.get("project", {}).get("project_name", ""),
                    "content_hash": entry.get("content_hash"),
                    "attempts": 0,
                    "last_error": "",
                    "updated_at": entry.get("generated_at"),
                }
            return fresh

    # ----------------------------------------------------------------- query
    def classify(self, tp: TrackedProject) -> str:
        """Return 'completed', 'outdated', 'failed', or 'new'."""
        cached = self.prompts.get(tp.project_id)
        if cached:
            return "completed" if cached.get("content_hash") == tp.content_hash else "outdated"
        state = self.progress["projects"].get(tp.project_id)
        if state and state.get("status") == STATUS_FAILED and state.get("content_hash") == tp.content_hash:
            return "failed"
        return "new"

    def cached_entry(self, tp: TrackedProject) -> dict | None:
        cached = self.prompts.get(tp.project_id)
        if cached and cached.get("content_hash") == tp.content_hash:
            return cached
        return None

    def generated_under_other_rules(self, tp: TrackedProject, rules_sha256: str) -> bool:
        cached = self.cached_entry(tp)
        return cached is not None and cached.get("rules_sha256") != rules_sha256

    def valid_entries(self, tracked: list[TrackedProject]) -> list[tuple[TrackedProject, dict]]:
        return [(tp, e) for tp in tracked if (e := self.cached_entry(tp)) is not None]

    # ---------------------------------------------------------------- update
    def _set_state(self, tp: TrackedProject, status: str, error: str = "", attempts_delta: int = 0) -> None:
        previous = self.progress["projects"].get(tp.project_id, {})
        attempts = previous.get("attempts", 0) if previous.get("content_hash") == tp.content_hash else 0
        self.progress["projects"][tp.project_id] = {
            "status": status,
            "excel_row": tp.project.excel_row,
            "project_name": tp.name,
            "content_hash": tp.content_hash,
            "attempts": attempts + attempts_delta,
            "last_error": error,
            "updated_at": now_iso(),
        }

    def mark_outdated(self, tp: TrackedProject) -> None:
        cached = self.prompts.get(tp.project_id)
        if cached and not cached.get("outdated"):
            cached["outdated"] = True
            self.save_prompts()
        self._set_state(tp, STATUS_OUTDATED, "Source information changed in Excel; regenerating")
        self.save_progress()

    def record_success(
        self, tp: TrackedProject, result, template_version: str, rules_sha256: str
    ) -> None:
        self.prompts[tp.project_id] = {
            "project_id": tp.project_id,
            "content_hash": tp.content_hash,
            "excel_row": tp.project.excel_row,
            "project": tp.project.fields(),
            "prompt": result.text,
            "model": result.model,
            "usage": result.usage,
            "request_id": result.request_id,
            "template_version": template_version,
            "rules_sha256": rules_sha256,
            "generated_at": now_iso(),
        }
        self.save_prompts()  # cache first: the paid-for result is never lost
        self._set_state(tp, STATUS_SUCCESS, attempts_delta=result.attempts)
        totals = self.progress["usage_totals"]
        for key, value in result.usage.items():
            totals[key] = totals.get(key, 0) + value
        totals["requests"] = totals.get("requests", 0) + 1
        self.save_progress()

    def record_failure(self, tp: TrackedProject, error: str, attempts: int = 1) -> None:
        self._set_state(tp, STATUS_FAILED, error=error, attempts_delta=attempts)
        self.save_progress()

    def record_run(self, summary: dict) -> None:
        self.progress["last_run"] = summary
        self.save_progress()

    def save_prompts(self) -> None:
        atomic_write_json(
            self.prompts_path,
            {"version": SCHEMA_VERSION, "updated_at": now_iso(), "prompts": self.prompts},
        )

    def save_progress(self) -> None:
        self.progress["version"] = SCHEMA_VERSION
        self.progress["updated_at"] = now_iso()
        atomic_write_json(self.progress_path, self.progress)
