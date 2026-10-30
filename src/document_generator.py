"""Build the Word document from cached prompts (python-docx).

The document is always rebuilt from scratch from the prompt cache, so it never
contains duplicate sections and never needs an API call to regenerate. It is
saved to a temporary file, re-opened to validate, then atomically swapped in.

Heading levels: Heading 1 = project title, Heading 2 = project sections
(both appear in the table of contents); the generated prompt's own Markdown
headings are mapped to Heading 3-6.
"""

from __future__ import annotations

import datetime as dt
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from src.progress_manager import replace_with_retry

TITLE = "Generated Claude Code Project Prompts"
BODY_FONT = "Calibri"
MONO_FONT = "Consolas"
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
MUTED = RGBColor(0x59, 0x59, 0x59)

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_INLINE = re.compile(r"(`[^`\n]+`|\*\*[^*\n]+\*\*|__[^_\n]+__|(?<![\w*])\*[^*\s][^*\n]*?\*(?!\*)|\[[^\]\n]+\]\([^)\s]+\))")
_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_BULLET = re.compile(r"^(\s*)[-*+]\s+(.*)$")
_NUMBERED = re.compile(r"^(\s*)(\d{1,3})[.)]\s+(.*)$")
_FENCE = re.compile(r"^\s*(```|~~~)")
_HR = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
_TABLE_SEP = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")


class DocumentError(Exception):
    """The Word document could not be written; cached prompts are unaffected."""


@dataclass(frozen=True)
class DocumentEntry:
    number: int
    project_name: str
    project_idea: str
    client_problem: str
    github_repo_name: str
    github_repo_description: str
    prompt: str


def clean_text(text: str) -> str:
    return _ILLEGAL_XML.sub("", text or "")


# ----------------------------------------------------------------- styles
def _configure_styles(doc) -> None:
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(10.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.1

    sizes = {1: 18, 2: 14, 3: 13, 4: 12, 5: 11, 6: 10.5}
    for level, size in sizes.items():
        style = styles[f"Heading {level}"]
        style.font.name = BODY_FONT
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = ACCENT
        style.paragraph_format.space_before = Pt(14 if level <= 2 else 10)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.keep_with_next = True

    if "Code Block" not in [s.name for s in styles]:
        code = styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
        code.base_style = styles["Normal"]
        code.font.name = MONO_FONT
        code.font.size = Pt(9)
        code.paragraph_format.space_before = Pt(4)
        code.paragraph_format.space_after = Pt(6)
        code.paragraph_format.line_spacing = 1.0
        code.paragraph_format.left_indent = Inches(0.15)
        p_pr = code.element.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F2F4F7")
        p_pr.append(shd)

    if "Inline Code" not in [s.name for s in styles]:
        inline = styles.add_style("Inline Code", WD_STYLE_TYPE.CHARACTER)
        inline.font.name = MONO_FONT
        inline.font.size = Pt(9.5)
        inline.font.color.rgb = RGBColor(0x8B, 0x1A, 0x3A)


def _configure_page(doc) -> None:
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)


def _add_field(paragraph, instruction: str, placeholder: str = "") -> None:
    """Insert a Word field (e.g. TOC, PAGE) that Word computes when updated."""
    def fld(kind):
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
        return el

    run = paragraph.add_run()
    run._r.append(fld("begin"))
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    run._r.append(instr)
    run._r.append(fld("separate"))
    shown = paragraph.add_run(placeholder)
    shown.italic = bool(placeholder)
    end = paragraph.add_run()
    end._r.append(fld("end"))


def _request_field_update(doc) -> None:
    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")


def _add_footer(doc) -> None:
    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    label = footer.add_run(f"{TITLE}  |  Page ")
    label.font.size = Pt(8.5)
    label.font.color.rgb = MUTED
    _add_field(footer, "PAGE")


# ------------------------------------------------------------- inline md
def add_inline(paragraph, text: str, bold: bool = False, italic: bool = False) -> None:
    text = clean_text(text)
    pos = 0
    for match in _INLINE.finditer(text):
        if match.start() > pos:
            _run(paragraph, text[pos : match.start()], bold, italic)
        token = match.group(0)
        if token.startswith("`"):
            paragraph.add_run(token[1:-1], style="Inline Code")
        elif token.startswith("**") or token.startswith("__"):
            _run(paragraph, token[2:-2], True, italic)
        elif token.startswith("*"):
            _run(paragraph, token[1:-1], bold, True)
        else:
            label, _, url = token[1:-1].partition("](")
            _run(paragraph, label if label == url else f"{label} ({url})", bold, italic)
        pos = match.end()
    if pos < len(text):
        _run(paragraph, text[pos:], bold, italic)


def _run(paragraph, text: str, bold: bool, italic: bool) -> None:
    run = paragraph.add_run(text)
    run.bold = bold or None
    run.italic = italic or None


# ------------------------------------------------------------ block md
def _split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [cell.strip() for cell in re.split(r"(?<!\\)\|", line)]


def _add_table(doc, rows: list[list[str]]) -> None:
    columns = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=columns)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, row in enumerate(rows):
        for c in range(columns):
            cell = table.cell(r, c)
            paragraph = cell.paragraphs[0]
            add_inline(paragraph, row[c] if c < len(row) else "", bold=(r == 0))
            for run in paragraph.runs:
                run.font.size = Pt(9.5)
            if r == 0:
                tc_pr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), "DCE6F1")
                tc_pr.append(shd)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_code(doc, lines: list[str]) -> None:
    paragraph = doc.add_paragraph(style="Code Block")
    for i, line in enumerate(lines):
        if i:
            paragraph.add_run().add_break()
        paragraph.add_run(clean_text(line.rstrip("\n")).replace("\t", "    "))


def _list_paragraph(doc, kind: str, level: int, text: str, number: str = ""):
    level = min(level, 2)
    if kind == "bullet":
        style = "List Bullet" if level == 0 else f"List Bullet {level + 1}"
        paragraph = doc.add_paragraph(style=style)
        text = re.sub(r"^\[ \]\s*", "☐ ", text)
        text = re.sub(r"^\[[xX]\]\s*", "☑ ", text)
        add_inline(paragraph, text)
    else:
        paragraph = doc.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.left_indent = Inches(0.35 + 0.3 * level)
        fmt.first_line_indent = Inches(-0.3)
        fmt.space_after = Pt(2)
        paragraph.add_run(f"{number}.\t")
        fmt.tab_stops.add_tab_stop(fmt.left_indent)
        add_inline(paragraph, text)
    paragraph.paragraph_format.space_after = Pt(2)
    return paragraph


def render_markdown(doc, markdown: str, heading_offset: int = 2) -> None:
    """Render Markdown (as produced by Claude) into structured Word content."""
    lines = clean_text(markdown).replace("\r\n", "\n").split("\n")
    paragraph_buffer: list[str] = []
    last_list = None  # paragraph of the last list item, for continuation lines
    i = 0

    def flush():
        nonlocal paragraph_buffer
        if paragraph_buffer:
            add_inline(doc.add_paragraph(), " ".join(s.strip() for s in paragraph_buffer))
            paragraph_buffer = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        fence_match = _FENCE.match(line)
        if fence_match:
            flush()
            fence = fence_match.group(1)
            code, i = [], i + 1
            while i < len(lines) and not lines[i].strip().startswith(fence):
                code.append(lines[i])
                i += 1
            _add_code(doc, code)
            last_list = None
            i += 1
            continue

        if not stripped:
            flush()
            last_list = None if (i + 1 >= len(lines) or not lines[i + 1].startswith(" ")) else last_list
            i += 1
            continue

        heading = _HEADING.match(line)
        if heading:
            flush()
            level = min(len(heading.group(1)) + heading_offset, 6)
            add_inline(doc.add_paragraph(style=f"Heading {level}"), heading.group(2))
            last_list = None
            i += 1
            continue

        if _HR.match(line):
            flush()
            last_list = None
            i += 1
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and _TABLE_SEP.match(lines[i + 1]):
            flush()
            rows = [_split_row(line)]
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(_split_row(lines[i]))
                i += 1
            _add_table(doc, rows)
            last_list = None
            continue

        bullet = _BULLET.match(line)
        if bullet:
            flush()
            level = len(bullet.group(1).replace("\t", "  ")) // 2
            last_list = _list_paragraph(doc, "bullet", level, bullet.group(2))
            i += 1
            continue

        numbered = _NUMBERED.match(line)
        if numbered:
            flush()
            level = len(numbered.group(1).replace("\t", "  ")) // 3
            last_list = _list_paragraph(doc, "number", level, numbered.group(3), numbered.group(2))
            i += 1
            continue

        if stripped.startswith(">"):
            flush()
            quote = doc.add_paragraph(style="Quote")
            add_inline(quote, stripped.lstrip(">").strip())
            last_list = None
            i += 1
            continue

        if last_list is not None and line.startswith(" "):
            add_inline(last_list, " " + stripped)
            i += 1
            continue

        last_list = None
        paragraph_buffer.append(line)
        i += 1
    flush()


# ------------------------------------------------------------- document
def _title_page(doc, source_name: str, count: int, generated: dt.datetime) -> None:
    for _ in range(6):
        doc.add_paragraph()
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(TITLE)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f"Generated automatically from {source_name}")
    run.font.size = Pt(14)
    run.font.color.rgb = MUTED
    doc.add_paragraph()
    for label, value in (
        ("Generation date", generated.strftime("%d %B %Y, %H:%M")),
        ("Successfully generated projects", str(count)),
    ):
        line = doc.add_paragraph()
        line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        line.add_run(f"{label}: ").bold = True
        line.add_run(value)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _table_of_contents(doc) -> None:
    heading = doc.add_paragraph()
    run = heading.add_run("Table of Contents")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = ACCENT
    note = doc.add_paragraph()
    hint = note.add_run(
        "If the table below is empty or out of date, right-click it in Microsoft Word "
        "and choose “Update Field” (or press F9)."
    )
    hint.italic = True
    hint.font.size = Pt(9)
    hint.font.color.rgb = MUTED
    _add_field(doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u', "Table of contents - update field to display.")


def _labelled(doc, label: str, value: str) -> None:
    label_p = doc.add_paragraph()
    label_p.paragraph_format.space_after = Pt(0)
    label_p.paragraph_format.keep_with_next = True
    run = label_p.add_run(f"{label}:")
    run.bold = True
    run.font.color.rgb = ACCENT
    value_p = doc.add_paragraph()
    value_p.paragraph_format.left_indent = Inches(0.2)
    value_p.paragraph_format.space_after = Pt(6)
    lines = clean_text(value).split("\n") if value else ["Not provided"]
    for i, line in enumerate(lines):
        if i:
            value_p.add_run().add_break()
        r = value_p.add_run(line)
        if not value:
            r.italic = True
            r.font.color.rgb = MUTED


def _project_section(doc, entry: DocumentEntry) -> None:
    doc.add_paragraph(style="Heading 1").add_run(
        clean_text(f"PROJECT {entry.number:03d} — {entry.project_name}")
    )
    doc.add_paragraph(style="Heading 2").add_run("Project Details")
    _labelled(doc, "Project Name", entry.project_name)
    if entry.project_idea:  # optional column; omitted when the workbook has none
        _labelled(doc, "Project Idea", entry.project_idea)
    _labelled(doc, "Client Problem Solved", entry.client_problem)
    _labelled(doc, "GitHub Repository", entry.github_repo_name)
    _labelled(doc, "GitHub Repository Description", entry.github_repo_description)
    doc.add_paragraph(style="Heading 2").add_run("CLAUDE CODE PROMPT")
    render_markdown(doc, entry.prompt, heading_offset=2)


def build_document(entries: list[DocumentEntry], source_name: str, generated: dt.datetime | None = None):
    doc = Document()
    _configure_styles(doc)
    _configure_page(doc)
    _add_footer(doc)
    doc.core_properties.title = TITLE
    doc.core_properties.subject = f"Generated automatically from {source_name}"
    _title_page(doc, source_name, len(entries), generated or dt.datetime.now())
    _table_of_contents(doc)
    if not entries:
        doc.add_paragraph().add_run("No prompts have been generated yet.").italic = True
    for entry in entries:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        _project_section(doc, entry)
    _request_field_update(doc)
    return doc


def count_project_headings(path: Path) -> int:
    doc = Document(str(path))
    return sum(
        1 for p in doc.paragraphs
        if p.style is not None and p.style.name == "Heading 1" and p.text.startswith("PROJECT ")
    )


def save_document(entries: list[DocumentEntry], output_path: Path, source_name: str) -> Path:
    """Build, validate, and atomically save the document. Raises DocumentError."""
    try:
        doc = build_document(entries, source_name)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=f".{output_path.stem}.", suffix=".docx", dir=output_path.parent, delete=False
        ) as handle:
            tmp = Path(handle.name)
        try:
            doc.save(str(tmp))
            found = count_project_headings(tmp)
            if found != len(entries):
                raise DocumentError(
                    f"Validation failed: expected {len(entries)} project sections, found {found}."
                )
            replace_with_retry(tmp, output_path)
        finally:
            tmp.unlink(missing_ok=True)
    except PermissionError as exc:
        raise DocumentError(
            f"Could not write {output_path.name}: the file is locked. Close it in Microsoft Word "
            "and run `python main.py --rebuild-document` (no API calls are needed)."
        ) from exc
    except DocumentError:
        raise
    except Exception as exc:  # python-docx/lxml can raise many types
        raise DocumentError(f"Word document generation failed: {type(exc).__name__}: {exc}") from exc
    return output_path
