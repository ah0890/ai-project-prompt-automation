import hashlib

import pytest
from openpyxl import Workbook

from src.excel_reader import WorkbookError, clean_cell, normalize_header, read_projects
from src.progress_manager import track_projects
from tests.conftest import HEADERS, make_workbook, project_row


def test_reads_valid_rows_with_original_row_numbers(tmp_path):
    path = make_workbook(tmp_path / "wb.xlsx", [project_row(1), project_row(2)])
    data = read_projects(path)
    assert data.sheet_name == "Projects"
    assert [p.excel_row for p in data.projects] == [2, 3]
    first = data.projects[0]
    assert first.project_name == "Project 1"
    assert first.github_repo_name == "project-1"
    assert first.client_problem == "Clients waste time on task 1"


@pytest.mark.parametrize(
    "variant",
    ["Project Name", "project_name", "Project_Name", "PROJECT NAME", "project name", "  Project   Name: "],
)
def test_header_normalization_variants(variant):
    assert normalize_header(variant) == "project name"


def test_alternative_column_names(tmp_path):
    headers = ["Idea", "Name", "Problem Solved", "Repository Name", "Repo Description"]
    path = make_workbook(tmp_path / "wb.xlsx", [project_row(1)], headers=headers)
    project = read_projects(path).projects[0]
    assert project.project_idea.startswith("Idea 1")
    assert project.github_repo_description == "Repository for project 1"


def test_missing_required_column_stops_with_clear_message(tmp_path):
    headers = HEADERS[:-1]  # no GitHub Repo Description
    path = make_workbook(tmp_path / "wb.xlsx", [project_row(1)[:-1]], headers=headers)
    with pytest.raises(WorkbookError) as exc:
        read_projects(path)
    message = str(exc.value)
    assert "GitHub Repo Description" in message
    assert "Projects" in message


def test_ambiguous_columns_are_rejected(tmp_path):
    headers = HEADERS + ["Repo Name"]  # two columns map to GitHub Repo Name
    path = make_workbook(tmp_path / "wb.xlsx", [project_row(1) + ("x",)], headers=headers)
    with pytest.raises(WorkbookError, match="ambiguous"):
        read_projects(path)


def test_empty_rows_and_missing_names(tmp_path):
    rows = [
        project_row(1),
        (None, None, None, None, None),
        ("   ", "", None, " ", None),
        ("An idea", None, "problem", "repo", "desc"),
        project_row(2),
    ]
    data = read_projects(make_workbook(tmp_path / "wb.xlsx", rows))
    assert [p.project_name for p in data.projects] == ["Project 1", "Project 2"]
    assert data.empty_rows == 2
    assert [(s.excel_row, s.reason) for s in data.skipped] == [(5, "Missing Project Name")]


def test_blank_optional_fields_are_kept_blank(tmp_path):
    rows = [(None, "Only Name", None, None, None)]
    project = read_projects(make_workbook(tmp_path / "wb.xlsx", rows)).projects[0]
    assert project.project_name == "Only Name"
    assert project.project_idea == project.github_repo_name == project.github_repo_description == ""


def test_whitespace_and_cell_types():
    assert clean_cell("  hello \t world  ") == "hello world"
    assert clean_cell("line 1  \r\n  line 2 ") == "line 1\nline 2"
    assert clean_cell(42.0) == "42"
    assert clean_cell(3.5) == "3.5"
    assert clean_cell(True) == "TRUE"
    assert clean_cell(None) == ""


def test_formula_cells_use_cached_value_without_crashing(tmp_path):
    path = tmp_path / "wb.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    ws.append(["idea", "Formula Project", "problem", '=LOWER("REPO")', "desc"])
    wb.save(path)
    project = read_projects(path).projects[0]
    # openpyxl-written files have no cached formula result, so the value is blank.
    assert project.github_repo_name == ""


def test_header_row_not_on_first_row(tmp_path):
    path = tmp_path / "wb.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Project Master List"])
    ws.append([])
    ws.append(HEADERS)
    ws.append(project_row(1))
    wb.save(path)
    data = read_projects(path)
    assert data.header_row == 3
    assert data.projects[0].excel_row == 4


def test_multiple_valid_sheets_require_selection(tmp_path):
    path = make_workbook(
        tmp_path / "wb.xlsx", [project_row(1)], sheets={"Archive": (HEADERS, [project_row(9)])}
    )
    with pytest.raises(WorkbookError, match="EXCEL_SHEET_NAME"):
        read_projects(path)
    assert read_projects(path, "Archive").projects[0].project_name == "Project 9"


def test_single_valid_sheet_among_others_is_auto_selected(tmp_path):
    path = make_workbook(
        tmp_path / "wb.xlsx", [project_row(1)], sheets={"Notes": (["Note"], [["hello"]])}
    )
    assert read_projects(path).sheet_name == "Projects"


def test_unknown_sheet_name(tmp_path):
    path = make_workbook(tmp_path / "wb.xlsx", [project_row(1)])
    with pytest.raises(WorkbookError, match="Available worksheets: Projects"):
        read_projects(path, "Nope")


def test_missing_workbook(tmp_path):
    with pytest.raises(WorkbookError, match="not found"):
        read_projects(tmp_path / "missing.xlsx")


def test_corrupt_workbook(tmp_path):
    path = tmp_path / "bad.xlsx"
    path.write_bytes(b"not a zip file")
    with pytest.raises(WorkbookError, match="Could not open"):
        read_projects(path)


def test_duplicate_projects(tmp_path):
    same_name_different_content = list(project_row(1))
    same_name_different_content[0] = "A different idea"
    rows = [project_row(1), project_row(1), tuple(same_name_different_content)]
    data = read_projects(make_workbook(tmp_path / "wb.xlsx", rows))
    tracked, skipped = track_projects(data.projects)
    assert [tp.project.excel_row for tp in tracked] == [2, 4]
    assert tracked[0].project_id != tracked[1].project_id
    assert tracked[1].project_id.endswith("-2")
    assert skipped[0].excel_row == 3 and "Duplicate of Excel row 2" in skipped[0].reason


def test_workbook_is_not_modified(tmp_path):
    path = make_workbook(tmp_path / "wb.xlsx", [project_row(1)])
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    mtime = path.stat().st_mtime_ns
    read_projects(path)
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before
    assert path.stat().st_mtime_ns == mtime


def test_project_idea_column_is_optional(tmp_path):
    # Mirrors the owner's real workbook: no Project Idea column, plus extra columns.
    headers = ["Project ID", "Project Name", "Client Problem Solved", "GitHub Repository Name",
               "GitHub Repository Description", "GitHub Repository URL", "Project Status"]
    rows = [("PRJ-001", "Supplier Reconciler", "Manual ticking", "supplier-reconciler", "Matches bills", None, "Planned")]
    data = read_projects(make_workbook(tmp_path / "wb.xlsx", rows, headers=headers))
    project = data.projects[0]
    assert project.project_name == "Supplier Reconciler"
    assert project.project_idea == ""
    assert project.github_repo_name == "supplier-reconciler"
