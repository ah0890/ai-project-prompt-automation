# context.md — Project Handoff

> Read this first in every new session. Then check `git status`, `git log --oneline -5`, and `git remote -v` before changing anything. Never store secrets here.

## Project objective
A local Python CLI that reads project ideas from `Project_Master_List.xlsx`, sends **one project per request** to the Anthropic Claude API, and writes the generated Claude Code development prompts into `output/Generated_Claude_Code_Project_Prompts.docx`. It must survive interruptions, never regenerate completed work unnecessarily, and never modify the workbook.

## Intended users
The repository owner, who maintains a list of software project ideas and wants implementation-ready Claude Code prompts for each.

## Current status (2026-10-31, v1.1.0)
- `prompt-rules.md` (owner-supplied, 294 lines) is the authoritative rule source. It is validated before any API call, embedded verbatim in every system prompt, and each response is compliance-checked. Output structure follows rules section 14.
- Real workbook `Projects_Master_List.xlsx`: sheet `Projects`, 44 projects (PRJ-001 to PRJ-044, all Planned), **no Project Idea column**. The owner confirmed Project Name (with client problem and repo fields) describes each project, so Project Idea is optional. The Guide sheet is ignored. The dry run passes.

## Status at v1.0.0
- **Implementation:** complete for all specified features (milestones 1–8 below).
- **Current milestone:** done, pending live API verification and real-workbook validation.
- **Verification:** 78 automated tests pass (all mocked). mypy reports no issues. pip-audit found no known vulnerabilities.
- **Not yet verified:**
  - **Live API generation:** no API key was available in the build session.
  - **Visual Word rendering:** Word COM automation is not registered on the dev machine. DOCX structure was verified with python-docx only.
- **Source rules file:** superseded. `prompt-rules.md` is now loaded at runtime.

## Architecture
```
main.py                     CLI + orchestration (Runner): config -> workbook -> track -> select -> generate -> cache -> docx -> summary
src/config.py               .env loading (python-dotenv), typed validation, paths resolved from project root
src/excel_reader.py         openpyxl read-only/data-only; header detection (first 20 rows), alias matching, sheet selection
src/prompt_template.py      load_prompt_rules + build_system_prompt(rules) (static per run, cacheable) + user message + check_compliance
src/claude_generator.py     streamed messages request; own retry/backoff (SDK retries off); response validation
src/progress_manager.py     project IDs, content hashes, generated_prompts.json + progress.json, atomic writes
src/document_generator.py   python-docx builder + Markdown->DOCX renderer; atomic save with validation
src/logger.py               CSV processing log (append+fsync), app.log, secret redaction, CSV-injection guard
tests/                      pytest; conftest.py has FakeClient (mocks messages.stream / beta.messages.stream)
```

## Key technical decisions
| Decision | Rationale |
| --- | --- |
| `project_id` = hash(normalized Project Name + Repo Name); `content_hash` = hash of all 5 fields (whitespace-normalized) | Row insertions and reordering don't trigger paid regeneration; edits do. Same identity with different content gets a `-2` suffix; identical rows are skipped as duplicates. |
| Prompt cache written **before** progress/doc | A paid result is never lost. `progress.json` can be rebuilt from the cache; the cache is never auto-overwritten when corrupt. |
| Word doc always rebuilt from cache | No duplicates; `--rebuild-document` needs no API. Outdated (changed-source) prompts stay in the cache but are left out of the doc. |
| SDK `max_retries=0`, custom loop | Exact `MAX_RETRIES`/`RETRY_BASE_DELAY` semantics; honours `retry-after`. |
| 401/403/404/402 = `FatalAPIError` stops the run | These fail identically for every project. |
| `stop_reason` `max_tokens`/`refusal`, empty, or <1500 chars are failures | Never cache truncated or bogus output. Empty/short responses are retryable. |
| Streaming (`messages.stream` + `get_final_message`) | Long outputs with `MAX_TOKENS=32000` would risk HTTP timeouts without it. |
| Effort default `medium`; `ANTHROPIC_EFFORT=default` omits it | Opus 5.5's default is medium; older models reject the param. |
| Server-side refusal fallback (`fallbacks: "default"`, beta `server-side-fallback-2026-07-01`) only for Opus 5.5/Opus 5/Sonnet 5.5/Fable 5.1 | Other models reject it. Toggle with `ENABLE_REFUSAL_FALLBACK`. |
| Default run = new + changed + previously failed; `--retry-failed` = only failed | Failed projects are always retried later; the flag narrows a run. |
| Numbered lists rendered as literal numbers with hanging indent | Avoids python-docx's continuous `List Number` numbering across the document. |
| `output/`, `logs/`, `data/`, and the workbook are git-ignored | Private content; owner can un-ignore. |

## Important files
- Config template: `.env.example` (blank placeholders). The real `.env` is git-ignored.
- Runtime state: `data/generated_prompts.json` (source of truth), `data/progress.json`, `logs/processing_log.csv`, `logs/app.log`.

## Configuration requirements
`ANTHROPIC_API_KEY` and `ANTHROPIC_MODEL` are required (recommended model: `claude-opus-5-5`). All optional settings and defaults are listed in README.md under "All settings".

## Commands
```powershell
python -m venv .venv; .\.venv\Scripts\activate; pip install -r requirements.txt
python -m pytest -q                       # mocked tests, no key needed
python main.py --dry-run                  # validate workbook/config
python main.py --limit 1                  # first live test (costs money)
python main.py                            # full run
python main.py --rebuild-document         # doc from cache, no API
```
Dev-only tools installed in `.venv` but not in requirements: `mypy`, `types-openpyxl`, `pip-audit`.

## Completed milestones
1. Workspace inspection and architecture
2. Configuration and Excel reader
3. Claude API integration (mock-verified)
4. Persistence and resume
5. Word document generation
6. Logging and CLI
7. Testing and security review
8. Documentation and verification

## Tests actually performed
- `pytest`: 78 passed. The suites are excel_reader, prompt_generator, progress_manager, document_generator, and main (end-to-end with temp workbooks and a FakeClient).
- Simulated end-to-end runs: crash mid-request then resume; network loss with retry on rerun; auth failure stops the run; Word save failure, then `--rebuild-document` with no key; changed and new projects; `--limit`; `--force`; `--dry-run`; secrets absent from logs and data; workbook SHA-256 unchanged.
- Real CLI: `--help`, plus a missing-key run (exit 2 with setup instructions).
- `mypy src main.py`: no issues. `pip-audit -r requirements.txt`: no known vulnerabilities.

## Pending work / next actions
1. Fix `ANTHROPIC_MODEL` in `.env` to an exact ID (currently a display name, `Opus-5.5`; use `claude-opus-5-5`).
2. Owner creates `.env`, then runs `python main.py --limit 1` and reviews prompt quality and token usage in the DOCX (live verification).
3. Tune `ANTHROPIC_EFFORT`/`MAX_TOKENS` based on that run. Bump `TEMPLATE_VERSION` if the template changes.
4. Commit milestones once the owner authorizes (see Git status).

## Known limitations / risks
- System prompt is about 2.6k tokens, possibly below the model's minimum cacheable prefix. If so, caching silently doesn't apply (harmless; visible in the summary).
- The Markdown converter covers common constructs only. Exotic Markdown renders as plain text.
- The TOC is a Word field; Word fills it in when the document is opened or the field is updated.
- If `--force` names a row, numbering uses Excel row numbers, not project numbers.

## Outstanding questions
- Git author email is `Ahaulakh9800@gmail.com` (repo config). The owner should confirm this is the intended identity before the first commit.
- Should the workbook and/or `output/` be versioned in Git? Both are currently ignored.

## Git status
- Branch `main`, remote `origin` = `https://github.com/ah0890/ai-project-prompt-automation.git`.
- Commits are authored by the owner's configured identity only, with no AI co-author or attribution lines (owner's instruction). See `git log` for the latest hash and push status.
