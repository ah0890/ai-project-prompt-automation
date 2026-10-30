"""Read project rows from the Excel workbook (read-only; never modified).

The workbook is opened with openpyxl in read-only, data-only mode, so formula
cells yield their cached values and nothing is ever written back.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

# Logical field -> accepted (normalized) header spellings.
COLUMN_ALIASES: dict[str, set[str]] = {
    "project_idea": {"project idea", "project description", "idea", "project ideas"},
    "project_name": {"project name", "name", "project title"},
    "client_problem": {"client problem solved", "problem solved", "client problem"},
    "github_repo_name": {
        "github repo name", "github repository name", "repository name",
        "repo name", "github repo", "github repository",
    },
    "github_repo_description": {
        "github repo description", "github repository description",
        "repository description", "repo description",
    },
}

DISPLAY_NAMES = {
    "project_idea": "Project Idea",
    "project_name": "Project Name",
    "client_problem": "Client Problem Solved",
    "github_repo_name": "GitHub Repo Name",
    "github_repo_description": "GitHub Repo Description",
}

# Project Idea is optional: the owner's workbook describes each project through
# its name, client problem, and repository description instead.
OPTIONAL_FIELDS = {"project_idea"}
REQUIRED_FIELDS = [f for f in COLUMN_ALIASES if f not in OPTIONAL_FIELDS]

HEADER_SCAN_ROWS = 20


class WorkbookError(Exception):
    """Raised when the workbook cannot be used; the message is user-facing."""


@dataclass(frozen=True)
class Project:
    excel_row: int
    project_idea: str
    project_name: str
    client_problem: str
    github_repo_name: str
    github_repo_description: str

    def fields(self) -> dict[str, str]:
        return {
            "project_idea": self.project_idea,
            "project_name": self.project_name,
            "client_problem": self.client_problem,
            "github_repo_name": self.github_repo_name,
            "github_repo_description": self.github_repo_description,
        }


@dataclass(frozen=True)
class SkippedRow:
    excel_row: int
    reason: str
    project_name: str = ""


@dataclass
class SheetLayout:
    sheet_name: str
    header_row: int | None
    columns: dict[str, int] = field(default_factory=dict)  # field -> 0-based index
    missing: list[str] = field(default_factory=list)
    ambiguous: dict[str, list[str]] = field(default_factory=dict)

    @property
    def valid(self) -> bool:
        return self.header_row is not None and not self.missing and not self.ambiguous


@dataclass
class WorkbookData:
    path: Path
    sheet_name: str
    header_row: int
    projects: list[Project]
    skipped: list[SkippedRow]
    empty_rows: int


def normalize_header(value: object) -> str:
    text = clean_cell(value).lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.rstrip(":").strip()


def clean_cell(value: object) -> str:
    """Convert any cell value to clean text without inventing content."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t ]+", " ", line).strip() for line in text.split("\n")]
    return "\n".join(lines).strip()


def _match_headers(header_cells: list[object]) -> tuple[dict[str, int], dict[str, list[str]]]:
    found: dict[str, list[tuple[int, str]]] = {}
    for index, cell in enumerate(header_cells):
        normalized = normalize_header(cell)
        if not normalized:
            continue
        for field_name, aliases in COLUMN_ALIASES.items():
            if normalized in aliases:
                found.setdefault(field_name, []).append((index, clean_cell(cell)))
    columns = {f: hits[0][0] for f, hits in found.items() if len(hits) == 1}
    ambiguous = {f: [h[1] for h in hits] for f, hits in found.items() if len(hits) > 1}
    return columns, ambiguous


def inspect_sheet(ws) -> SheetLayout:
    """Find the header row (best match within the first rows) and map columns."""
    best: SheetLayout | None = None
    for row_number, row in enumerate(
        ws.iter_rows(min_row=1, max_row=HEADER_SCAN_ROWS, values_only=True), start=1
    ):
        columns, ambiguous = _match_headers(list(row))
        if not columns and not ambiguous:
            continue
        layout = SheetLayout(
            sheet_name=ws.title,
            header_row=row_number,
            columns=columns,
            missing=[f for f in REQUIRED_FIELDS if f not in columns and f not in ambiguous],
            ambiguous=ambiguous,
        )
        if best is None or len(layout.columns) > len(best.columns):
            best = layout
        if layout.valid:
            return layout
    return best or SheetLayout(ws.title, None, missing=list(REQUIRED_FIELDS))


def describe_layout_problem(layout: SheetLayout) -> str:
    if layout.header_row is None:
        return f"worksheet '{layout.sheet_name}': no recognizable header row in the first {HEADER_SCAN_ROWS} rows"
    parts = []
    if layout.missing:
        names = ", ".join(DISPLAY_NAMES[f] for f in layout.missing)
        parts.append(f"missing required column(s): {names}")
    for f, headers in layout.ambiguous.items():
        parts.append(
            f"ambiguous column for {DISPLAY_NAMES[f]}: {', '.join(repr(h) for h in headers)} "
            "(rename or remove all but one)"
        )
    return f"worksheet '{layout.sheet_name}' (header row {layout.header_row}): " + "; ".join(parts)


def _open(path: Path):
    if not path.exists():
        raise WorkbookError(
            f"Excel workbook not found: {path}\n"
            "Place Project_Master_List.xlsx in the project root, or set EXCEL_FILE in .env."
        )
    try:
        return load_workbook(path, read_only=True, data_only=True)
    except (InvalidFileException, BadZipFile, OSError, KeyError, ValueError) as exc:
        raise WorkbookError(f"Could not open workbook {path.name}: {exc}") from exc


def select_sheet(wb, sheet_name: str | None, path: Path) -> tuple[Any, SheetLayout]:
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise WorkbookError(
                f"Worksheet '{sheet_name}' (EXCEL_SHEET_NAME) not found in {path.name}. "
                f"Available worksheets: {', '.join(wb.sheetnames)}"
            )
        ws = wb[sheet_name]
        layout = inspect_sheet(ws)
        if not layout.valid:
            raise WorkbookError(f"Cannot use {path.name}, {describe_layout_problem(layout)}.")
        return ws, layout

    layouts = [(wb[name], inspect_sheet(wb[name])) for name in wb.sheetnames]
    valid = [(ws, layout) for ws, layout in layouts if layout.valid]
    if len(valid) == 1:
        return valid[0]
    if not valid:
        details = "\n  - ".join(describe_layout_problem(layout) for _, layout in layouts)
        raise WorkbookError(
            f"No worksheet in {path.name} contains all required columns.\n  - {details}"
        )
    names = ", ".join(f"'{layout.sheet_name}'" for _, layout in valid)
    raise WorkbookError(
        f"Multiple worksheets in {path.name} contain project data: {names}.\n"
        "Choose one by setting EXCEL_SHEET_NAME in .env or passing --sheet NAME."
    )


def read_projects(path: Path, sheet_name: str | None = None) -> WorkbookData:
    wb = _open(path)
    try:
        ws, layout = select_sheet(wb, sheet_name, path)
        header_row = layout.header_row
        assert header_row is not None  # guaranteed by layout.valid
        projects: list[Project] = []
        skipped: list[SkippedRow] = []
        empty_rows = 0
        width = max(layout.columns.values()) + 1
        for row_number, row in enumerate(
            ws.iter_rows(min_row=header_row + 1, values_only=True),
            start=header_row + 1,
        ):
            cells = list(row) + [None] * max(0, width - len(row))
            values = {f: "" for f in COLUMN_ALIASES}  # optional columns may be absent
            values.update({f: clean_cell(cells[i]) for f, i in layout.columns.items()})
            if not any(values.values()):
                empty_rows += 1
                continue
            if not values["project_name"]:
                skipped.append(SkippedRow(row_number, "Missing Project Name"))
                continue
            projects.append(Project(excel_row=row_number, **values))
        return WorkbookData(
            path=path,
            sheet_name=layout.sheet_name,
            header_row=header_row,
            projects=projects,
            skipped=skipped,
            empty_rows=empty_rows,
        )
    finally:
        wb.close()
