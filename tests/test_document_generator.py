import pytest
from docx import Document

from src.document_generator import (
    DocumentEntry,
    DocumentError,
    build_document,
    clean_text,
    save_document,
)
from tests.conftest import SAMPLE_PROMPT


def entry(n, prompt=SAMPLE_PROMPT, **overrides):
    values = dict(
        number=n,
        project_name=f"Project {n}",
        project_idea=f"Idea {n}",
        client_problem=f"Problem {n}",
        github_repo_name=f"repo-{n}",
        github_repo_description=f"Description {n}",
        prompt=prompt,
    )
    values.update(overrides)
    return DocumentEntry(**values)


def all_text(doc):
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def save_and_open(tmp_path, entries):
    path = save_document(entries, tmp_path / "out" / "Generated_Claude_Code_Project_Prompts.docx",
                         "Project_Master_List.xlsx")
    return Document(str(path))


def test_valid_docx_with_title_page_and_metadata(tmp_path):
    doc = save_and_open(tmp_path, [entry(1)])
    text = all_text(doc)
    assert "Generated Claude Code Project Prompts" in text
    assert "Generated automatically from Project_Master_List.xlsx" in text
    assert "Successfully generated projects: 1" in text
    for label in ["Project Name:", "Project Idea:", "Client Problem Solved:", "GitHub Repository:",
                  "GitHub Repository Description:", "CLAUDE CODE PROMPT"]:
        assert label in text
    for value in ["Idea 1", "Problem 1", "repo-1", "Description 1"]:
        assert value in text


def test_heading_structure_and_prompt_formatting(tmp_path):
    doc = save_and_open(tmp_path, [entry(1)])
    styles = [(p.style.name, p.text) for p in doc.paragraphs]
    assert ("Heading 1", "PROJECT 001 — Project 1") in styles
    assert ("Heading 2", "CLAUDE CODE PROMPT") in styles
    assert ("Heading 3", "Claude Code Prompt: Sample") in styles      # '#'  -> Heading 3
    assert ("Heading 4", "A. Selected Expert Personas") in styles     # '##' -> Heading 4
    assert any(s == "List Bullet" for s, _ in styles)
    assert any(s == "List Bullet 2" for s, _ in styles)               # nested bullet
    assert any(s == "Code Block" and "python main.py --help" in t for s, t in styles)
    assert any(t.startswith("1.\tRead input") for _, t in styles)
    assert len(doc.tables) == 1 and doc.tables[0].cell(1, 0).text == "Python"
    bold = [r.text for p in doc.paragraphs for r in p.runs if r.bold]
    assert "Python Automation Engineer" in bold
    inline_code = [r.text for p in doc.paragraphs for r in p.runs if r.style and r.style.name == "Inline Code"]
    assert "src/" in inline_code
    assert "link (https://example.com)" in all_text(doc)


def test_complete_prompt_content_is_preserved(tmp_path):
    doc = save_and_open(tmp_path, [entry(1)])
    text = all_text(doc)
    assert text.count("Additional implementation detail.") == 60


def test_multiple_sections_without_duplicates(tmp_path):
    entries = [entry(1), entry(2), entry(3)]
    path = tmp_path / "out" / "doc.docx"
    save_document(entries, path, "wb.xlsx")
    save_document(entries, path, "wb.xlsx")  # rebuilding never appends
    doc = Document(str(path))
    headings = [p.text for p in doc.paragraphs if p.style.name == "Heading 1"]
    assert headings == [f"PROJECT 00{i} — Project {i}" for i in (1, 2, 3)]


def test_table_of_contents_field_and_page_breaks(tmp_path):
    doc = save_and_open(tmp_path, [entry(1), entry(2)])
    xml = doc.element.xml
    assert 'TOC \\o "1-2" \\h \\z \\u' in xml
    assert xml.count('w:type="page"') >= 3  # title page, before project 1, before project 2


def test_empty_cache_produces_valid_document(tmp_path):
    doc = save_and_open(tmp_path, [])
    assert "No prompts have been generated yet." in all_text(doc)


def test_blank_metadata_shows_not_provided(tmp_path):
    doc = save_and_open(tmp_path, [entry(1, github_repo_description="")])
    assert "Not provided" in all_text(doc)


def test_illegal_xml_characters_are_removed(tmp_path):
    doc = save_and_open(tmp_path, [entry(1, prompt="# Title\n\nBad\x00char\x0b here " + "x" * 50)])
    assert "Badchar here" in all_text(doc)
    assert clean_text("a\x01b") == "ab"


def test_locked_output_reports_clear_error(tmp_path, monkeypatch):
    def locked(*_a, **_k):
        raise PermissionError("in use")

    monkeypatch.setattr("src.document_generator.replace_with_retry", locked)
    with pytest.raises(DocumentError, match="Close it in Microsoft Word"):
        save_document([entry(1)], tmp_path / "doc.docx", "wb.xlsx")
    assert not list(tmp_path.glob("*.docx"))  # no partial file left behind


def test_existing_document_survives_failed_rebuild(tmp_path, monkeypatch):
    path = tmp_path / "doc.docx"
    save_document([entry(1)], path, "wb.xlsx")
    original = path.read_bytes()

    def broken(*_a, **_k):
        raise RuntimeError("renderer crashed")

    monkeypatch.setattr("src.document_generator.render_markdown", broken)
    with pytest.raises(DocumentError, match="renderer crashed"):
        save_document([entry(1), entry(2)], path, "wb.xlsx")
    assert path.read_bytes() == original


def test_build_document_numbering_uses_entry_number():
    doc = build_document([entry(7)], "wb.xlsx")
    assert any(p.text.startswith("PROJECT 007") for p in doc.paragraphs)
