"""End-to-end workflow tests: temporary workbook + mocked Claude responses.

These are simulated runs; no live Anthropic API request is ever made.
"""

import csv
import hashlib
import json
from pathlib import Path

import anthropic
import pytest
from docx import Document

import main as app
from src.claude_generator import ClaudeGenerator
from src.document_generator import DocumentError
from tests.conftest import (
    FakeClient,
    api_error,
    connection_error,
    fake_message,
    make_workbook,
    project_row,
)


class Harness:
    def __init__(self, tmp_path: Path, env: dict):
        self.tmp = tmp_path
        self.env = env
        self.client = FakeClient()
        self.output: list[str] = []

    @property
    def workbook(self) -> Path:
        return Path(self.env["EXCEL_FILE"])

    def write_rows(self, rows):
        make_workbook(self.workbook, rows)

    def factory(self, cfg, rules):
        return ClaudeGenerator(cfg, rules, client=self.client, sleep=lambda _s: None)

    def run(self, *argv, outcomes=None):
        self.client.outcomes = list(outcomes or [])
        self.client.requests.clear()
        self.output = []
        return app.main(list(argv), environ=self.env, generator_factory=self.factory, out=self.output.append)

    @property
    def text(self) -> str:
        return "\n".join(self.output)

    def summary(self) -> dict:
        values = {}
        for line in self.output:
            if ": " in line and line.split(": ")[0] in {
                "Total Projects", "Successfully Generated", "Failed", "Skipped", "Not Yet Processed",
                "Previously Completed", "Newly Processed", "Failed This Run",
            }:
                key, value = line.split(": ", 1)
                values[key] = int(value)
        return values

    def csv_rows(self):
        path = self.tmp / "logs" / "processing_log.csv"
        with path.open(encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))

    def doc_headings(self):
        doc = Document(str(self.tmp / "output" / "Generated_Claude_Code_Project_Prompts.docx"))
        return [p.text for p in doc.paragraphs if p.style.name == "Heading 1"]

    def cache(self):
        return json.loads((self.tmp / "data" / "generated_prompts.json").read_text(encoding="utf-8"))["prompts"]


@pytest.fixture
def h(tmp_path, env):
    return Harness(tmp_path, env)


def test_full_workflow(h):
    h.write_rows([project_row(1), project_row(2), (None, None, None, None, None),
                  ("idea", None, "p", "r", "d"), project_row(3)])
    before = hashlib.sha256(h.workbook.read_bytes()).hexdigest()

    assert h.run() == 0
    assert len(h.client.requests) == 3
    assert h.summary() == {"Total Projects": 4, "Successfully Generated": 3, "Failed": 0, "Skipped": 1,
                           "Previously Completed": 0, "Newly Processed": 3}
    assert h.doc_headings() == ["PROJECT 001 — Project 1", "PROJECT 002 — Project 2",
                                "PROJECT 003 — Project 3"]
    statuses = [(r["Excel Row Number"], r["Status"]) for r in h.csv_rows()]
    assert statuses == [("5", "SKIPPED"), ("2", "SUCCESS"), ("3", "SUCCESS"), ("6", "SUCCESS")]
    assert h.csv_rows()[1]["Project Number"] == "001"
    assert hashlib.sha256(h.workbook.read_bytes()).hexdigest() == before  # workbook untouched
    progress = json.loads((h.tmp / "data" / "progress.json").read_text(encoding="utf-8"))
    assert progress["last_run"]["newly_generated"] == 3


def test_rerun_does_not_regenerate_completed_projects(h):
    h.write_rows([project_row(1), project_row(2)])
    assert h.run() == 0
    assert h.run() == 0
    assert h.client.requests == []
    assert "Nothing to process" in h.text
    s = h.summary()
    assert s["Previously Completed"] == 2 and s["Newly Processed"] == 0 and s["Successfully Generated"] == 2
    assert len(h.doc_headings()) == 2


def test_crash_during_request_then_resume(h):
    h.write_rows([project_row(i) for i in range(1, 6)])
    code = h.run(outcomes=[fake_message(), fake_message(), KeyboardInterrupt()])
    assert code == app.EXIT_INTERRUPTED
    assert len(h.cache()) == 2
    assert h.summary()["Not Yet Processed"] == 3

    assert h.run() == 0
    assert len(h.client.requests) == 3  # only the 3 unfinished projects
    s = h.summary()
    assert s["Previously Completed"] == 2 and s["Newly Processed"] == 3 and s["Successfully Generated"] == 5
    assert len(h.doc_headings()) == 5


def test_network_failure_continues_and_failed_project_retried_later(h):
    h.write_rows([project_row(1), project_row(2), project_row(3)])
    # MAX_RETRIES=1 -> 2 attempts per project; project 2 fails both.
    code = h.run(outcomes=[fake_message(), connection_error(), connection_error(), fake_message()])
    assert code == app.EXIT_FAILURES
    s = h.summary()
    assert s["Successfully Generated"] == 2 and s["Failed"] == 1
    failed = [r for r in h.csv_rows() if r["Status"] == "FAILED"]
    assert failed[0]["Project Name"] == "Project 2" and "connection" in failed[0]["Error Message"].lower()
    assert h.doc_headings() == ["PROJECT 001 — Project 1", "PROJECT 003 — Project 3"]

    assert h.run("--retry-failed") == 0
    assert len(h.client.requests) == 1
    assert h.summary()["Failed"] == 0
    assert len(h.doc_headings()) == 3


def test_failed_projects_also_retried_by_normal_run(h):
    h.write_rows([project_row(1)])
    h.run(outcomes=[api_error(anthropic.BadRequestError, 400)])
    assert h.run() == 0
    assert len(h.client.requests) == 1


def test_authentication_error_stops_run(h):
    h.write_rows([project_row(1), project_row(2), project_row(3)])
    code = h.run(outcomes=[api_error(anthropic.AuthenticationError, 401)])
    assert code == app.EXIT_FAILURES
    assert len(h.client.requests) == 1  # not retried, other projects not attempted
    assert "Authentication failed" in h.text
    assert h.summary()["Not Yet Processed"] == 2


def test_document_failure_keeps_prompt_and_rebuild_needs_no_api(h, monkeypatch):
    h.write_rows([project_row(1)])
    real_save = app.save_document

    def failing(*_a, **_k):
        raise DocumentError("simulated Word failure")

    monkeypatch.setattr(app, "save_document", failing)
    assert h.run() == app.EXIT_FAILURES
    assert len(h.cache()) == 1
    assert "[DOCUMENT]" in h.csv_rows()[-1]["Error Message"]

    monkeypatch.setattr(app, "save_document", real_save)
    h.env["ANTHROPIC_API_KEY"] = ""  # rebuild must not need the API
    assert h.run("--rebuild-document") == 0
    assert h.client.requests == []
    assert h.doc_headings() == ["PROJECT 001 — Project 1"]


def test_changed_project_is_regenerated_and_new_projects_added(h):
    h.write_rows([project_row(1), project_row(2)])
    h.run()
    changed = list(project_row(2))
    changed[2] = "A completely different client problem"
    h.write_rows([project_row(1), tuple(changed), project_row(3)])
    assert h.run() == 0
    assert len(h.client.requests) == 2  # changed project 2 + new project 3
    assert "Source information changed" in h.text
    s = h.summary()
    assert s["Previously Completed"] == 1 and s["Newly Processed"] == 2 and s["Successfully Generated"] == 3
    assert "A completely different client problem" in h.client.requests[0]["messages"][0]["content"]
    assert len(h.doc_headings()) == 3


def test_outdated_prompt_excluded_from_document_if_regeneration_fails(h):
    h.write_rows([project_row(1)])
    h.run()
    changed = list(project_row(1))
    changed[0] = "new idea"
    h.write_rows([tuple(changed)])
    h.run(outcomes=[api_error(anthropic.BadRequestError, 400)])
    assert h.doc_headings() == []  # stale prompt is not presented as current
    assert len(h.cache()) == 1  # but it is kept in the cache


def test_missing_api_key_explains_fix_and_makes_no_requests(h):
    h.write_rows([project_row(1)])
    h.env["ANTHROPIC_API_KEY"] = ""
    assert h.run() == app.EXIT_CONFIG
    assert "Create a file named .env" in h.text
    assert h.client.requests == []


def test_dry_run_makes_no_requests(h):
    h.write_rows([project_row(1), project_row(2)])
    assert h.run("--dry-run") == 0
    assert h.client.requests == []
    assert "Would process 2 project(s)" in h.text
    assert not (h.tmp / "data" / "generated_prompts.json").exists()


def test_limit_and_force(h):
    h.write_rows([project_row(1), project_row(2), project_row(3)])
    assert h.run("--limit", "2") == 0
    assert len(h.client.requests) == 2
    assert h.summary()["Not Yet Processed"] == 1

    assert h.run("--force", "2", "--limit", "1") == 0
    assert len(h.client.requests) == 1
    assert "Forced regeneration" in h.text
    assert "Project 1" in h.client.requests[0]["messages"][0]["content"]


def test_missing_workbook_and_columns(h):
    assert h.run() == app.EXIT_CONFIG
    assert "Excel workbook not found" in h.text
    make_workbook(h.workbook, [("a", "b")], headers=["Project Name", "Project Idea"])
    assert h.run() == app.EXIT_CONFIG
    assert "missing required column(s)" in h.text and "Client Problem Solved" in h.text


def test_secrets_never_written_to_outputs(h):
    h.write_rows([project_row(1), project_row(2)])
    key = h.env["ANTHROPIC_API_KEY"]
    h.run(outcomes=[fake_message(), api_error(anthropic.BadRequestError, 400)])
    for path in (h.tmp / "logs").iterdir():
        assert key not in path.read_text(encoding="utf-8")
    for path in (h.tmp / "data").iterdir():
        assert key not in path.read_text(encoding="utf-8")
    assert key not in h.text


def test_invalid_config_value(h):
    h.env["MAX_RETRIES"] = "lots"
    assert h.run() == app.EXIT_CONFIG
    assert "MAX_RETRIES must be a whole number" in h.text


def test_missing_rules_file_stops_before_any_request(h):
    h.write_rows([project_row(1)])
    Path(h.env["PROMPT_RULES_FILE"]).unlink()
    assert h.run() == app.EXIT_CONFIG
    assert "Prompt rules file not found" in h.text
    assert h.client.requests == []
    assert h.run("--dry-run") == app.EXIT_CONFIG  # dry run validates the rules too


def test_rules_are_reported_and_sent_with_every_request(h):
    h.write_rows([project_row(1), project_row(2)])
    assert h.run() == 0
    assert "Prompt rules:" in h.text and "LOADED" in h.text
    rules_text = Path(h.env["PROMPT_RULES_FILE"]).read_text(encoding="utf-8").strip()
    assert all(rules_text in r["system"][0]["text"] for r in h.client.requests)


def test_changed_rules_are_flagged_and_regenerated_only_on_request(h):
    h.write_rows([project_row(1), project_row(2)])
    h.run()
    rules_path = Path(h.env["PROMPT_RULES_FILE"])
    rules_path.write_text(rules_path.read_text(encoding="utf-8") + "\n## 2. New rule\nBe concise.\n", encoding="utf-8")

    assert h.run() == 0
    assert h.client.requests == []  # never spends money silently
    assert "Generated under a different prompt-rules.md: 2" in h.text

    assert h.run("--apply-new-rules", "--limit", "1") == 0
    assert len(h.client.requests) == 1
    assert "New rule" in h.client.requests[0]["system"][0]["text"]
    assert h.run() == 0
    assert "Generated under a different prompt-rules.md: 1" in h.text
