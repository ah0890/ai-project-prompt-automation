# Changelog

All notable changes to this project are recorded here. Only verified work is listed as completed.

## [1.1.0] - 2026-10-31

### Added
- Mandatory `prompt-rules.md` gate: validated before any API request (also in `--dry-run`), embedded verbatim in every request as the authoritative rules, with its SHA-256 stored per cached prompt.
- Post-generation compliance check against the rules' mandatory elements; one regeneration on failure.
- `--apply-new-rules` to regenerate prompts made under an older rules file; such prompts are otherwise only reported.
- `PROMPT_RULES_FILE` setting; workbook auto-detection for `Projects_Master_List.xlsx`.

### Changed
- Generated prompt structure now follows rules section 14 (Part 1 repository details, Part 2 with 16 sections). Template version 2.0.
- Project Idea column is optional; when absent, the project is described by its name, client problem, and repository description, and the Word document omits the Project Idea label.

### Security
- `.env.example` restored to blank placeholders; configured values live in git-ignored `.env`.

### Verification
- 88/88 tests passed (mocked). mypy reports no issues.
- Dry run against the real `Projects_Master_List.xlsx`: 44 projects detected, rules loaded, no API requests.
- Not yet verified: live API generation.

## [1.0.0] - 2026-10-31

### Added
- `main.py` CLI with `--dry-run`, `--retry-failed`, `--rebuild-document`, `--limit N`, `--force ROW...`, `--sheet NAME`; accurate run summary; exit codes 0/1/2/130.
- `src/config.py`: `.env` configuration with validation and clear missing-key guidance; secrets hidden from `repr`.
- `src/excel_reader.py`: read-only workbook access, header detection, column-name normalization and aliases, ambiguity detection, worksheet selection, row skipping with original row numbers.
- `src/prompt_template.py`: structured prompt template (sections A–R plus mandatory engineering standards); static, cacheable system prompt.
- `src/claude_generator.py`: streamed Anthropic requests, exponential backoff with `retry-after`, fatal-error handling, response validation (empty, short, truncated, refusal), optional server-side refusal fallback.
- `src/progress_manager.py`: stable project IDs, content-hash change detection, duplicate detection, atomic JSON persistence, corruption recovery.
- `src/document_generator.py`: Word document with title page, TOC field, page footer, project sections, and Markdown rendering (headings, lists, tables, code, quotes, inline formatting); atomic, validated save.
- `src/logger.py`: CSV processing log with fsync-per-row, app log, secret redaction, CSV formula-injection guard.
- Test suite (78 tests, mocked API) covering reader, generator, persistence, document, and end-to-end recovery scenarios.
- `README.md`, `context.md`, `CHANGELOG.md`, `.env.example`, `.gitignore`, `requirements.txt`.

### Changed
- `README.md`: replaced the initial three-line description with full documentation. The original text mentioned "automated research", which this tool does not perform.

### Verification
- 78/78 tests passed. mypy reports no issues. pip-audit found no known vulnerabilities.
- **Not yet verified:** live Claude API generation (no key available), real `Project_Master_List.xlsx` (not in workspace), visual rendering in Microsoft Word.
