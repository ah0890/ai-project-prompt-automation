"""Shared fixtures: temporary workbooks, configs, and a fake Anthropic client."""

from __future__ import annotations

import re
from pathlib import Path
from types import SimpleNamespace

import anthropic
import httpx2
import pytest
from openpyxl import Workbook

from src.config import Config
from src.prompt_template import PromptRules, load_prompt_rules

HEADERS = [
    "Project Idea",
    "Project Name",
    "Client Problem Solved",
    "GitHub Repo Name",
    "GitHub Repo Description",
]


def make_workbook(path: Path, rows, headers=HEADERS, sheets: dict | None = None) -> Path:
    """Create an .xlsx; `sheets` maps extra sheet names to (headers, rows)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Projects"
    ws.append(headers)
    for row in rows:
        ws.append(list(row))
    for name, (extra_headers, extra_rows) in (sheets or {}).items():
        extra = wb.create_sheet(name)
        extra.append(extra_headers)
        for row in extra_rows:
            extra.append(list(row))
    wb.save(path)
    return path


def project_row(n: int):
    return (
        f"Idea {n}: a tool that automates task {n}",
        f"Project {n}",
        f"Clients waste time on task {n}",
        f"project-{n}",
        f"Repository for project {n}",
    )


RULES_TEXT = """# Test Prompt Rules

## 1. Repository Details
Always include the GitHub repository name and description.
"""

SAMPLE_PROMPT = """# Claude Code Prompt: Sample

## Part 1 - GitHub Repository Details
- **Repository Name:** {repo}

## Part 2 - Complete Copy-Paste-Ready Coding Prompt
Maintain README.md, context.md and CHANGELOG.md. Verify the Git author identity.
Work milestone by milestone. Wait for DONE — GO TO LINKEDIN before LinkedIn work.

## A. Selected Expert Personas
- **Python Automation Engineer** - builds the core pipeline.
- Backend Engineer - designs `src/` modules.

## H. System Architecture
1. Read input
2. Process data
   - nested detail

```bash
python main.py --help
```

| Technology | Purpose |
|---|---|
| Python | Core |

Plain paragraph with *italic* and a [link](https://example.com).
""" + ("Additional implementation detail. " * 60)


def fake_message(text=SAMPLE_PROMPT, stop_reason="end_turn", usage=None):
    usage = usage or {"input_tokens": 100, "output_tokens": 2000,
                      "cache_creation_input_tokens": 0, "cache_read_input_tokens": 3000}
    return SimpleNamespace(
        content=[SimpleNamespace(type="text", text=text)] if text is not None else [],
        stop_reason=stop_reason,
        stop_details=None,
        model="claude-test-model",
        usage=SimpleNamespace(**usage),
        _request_id="req_test",
    )


def _response(status: int, headers=None):
    request = httpx2.Request("POST", "https://api.anthropic.com/v1/messages")
    return httpx2.Response(status, request=request, headers=headers or {})


def api_error(cls, status: int, headers=None):
    return cls(f"error {status}", response=_response(status, headers), body=None)


def connection_error():
    return anthropic.APIConnectionError(request=httpx2.Request("POST", "https://api.anthropic.com"))


def timeout_error():
    return anthropic.APITimeoutError(request=httpx2.Request("POST", "https://api.anthropic.com"))


class _Stream:
    def __init__(self, outcome):
        self.outcome = outcome

    def __enter__(self):
        if isinstance(self.outcome, BaseException):
            raise self.outcome
        return self

    def __exit__(self, *exc):
        return False

    def get_final_message(self):
        return self.outcome


class FakeMessages:
    def __init__(self, owner):
        self.owner = owner

    def stream(self, **kwargs):
        self.owner.requests.append(kwargs)
        outcome = self.owner.outcomes.pop(0) if self.owner.outcomes else self.owner.default()
        # Fill the "{repo}" placeholder with the requested repository name so
        # sample replies are rules-compliant for whichever project was sent.
        match = re.search(r"<github_repo_name>(.*?)</github_repo_name>", kwargs["messages"][0]["content"])
        for block in getattr(outcome, "content", None) or []:
            if getattr(block, "text", None):
                block.text = block.text.replace("{repo}", match.group(1) if match else "")
        return _Stream(outcome)


class FakeClient:
    """Mimics client.messages.stream / client.beta.messages.stream.

    `outcomes` is consumed in order; each item is a message or an exception.
    """

    def __init__(self, outcomes=None, default=fake_message):
        self.outcomes = list(outcomes or [])
        self.default = default
        self.requests: list[dict] = []
        self.messages = FakeMessages(self)
        self.beta = SimpleNamespace(messages=FakeMessages(self))


@pytest.fixture
def rules(tmp_path) -> PromptRules:
    path = tmp_path / "prompt-rules.md"
    path.write_text(RULES_TEXT, encoding="utf-8")
    return load_prompt_rules(path)


@pytest.fixture
def config(tmp_path) -> Config:
    return Config(
        api_key="sk-ant-test-key-1234567890",
        model="claude-test-model",
        max_retries=3,
        retry_base_delay=2.0,
        excel_file=tmp_path / "Project_Master_List.xlsx",
        output_dir=tmp_path / "output",
        log_dir=tmp_path / "logs",
        data_dir=tmp_path / "data",
    )


@pytest.fixture
def env(tmp_path) -> dict:
    """Environment for main.main(); points every path into tmp_path."""
    return {
        "ANTHROPIC_API_KEY": "sk-ant-test-key-1234567890",
        "ANTHROPIC_MODEL": "claude-test-model",
        "EXCEL_FILE": str(tmp_path / "Project_Master_List.xlsx"),
        "OUTPUT_DIR": str(tmp_path / "output"),
        "LOG_DIR": str(tmp_path / "logs"),
        "DATA_DIR": str(tmp_path / "data"),
        "RETRY_BASE_DELAY": "0",
        "MAX_RETRIES": "1",
        "PROMPT_RULES_FILE": str(_write_rules(tmp_path)),
    }


def _write_rules(tmp_path: Path) -> Path:
    path = tmp_path / "prompt-rules.md"
    path.write_text(RULES_TEXT, encoding="utf-8")
    return path
