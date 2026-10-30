# ai-project-prompt-automation

A Python automation tool (working title: **Claude Code Prompt Generator**) that reads software project ideas from an Excel workbook and uses the Anthropic Claude API to generate detailed, implementation-ready **Claude Code** development prompts, collected in one professionally formatted Word document.

## Project overview

You keep a list of projects in `Projects_Master_List.xlsx` (project name, client problem, GitHub repo name and description; a project idea column is optional). Turning each idea into a thorough specification for Claude Code takes a long time by hand. This tool:

1. loads and validates `prompt-rules.md`, the authoritative rules for every generated prompt, and stops before any API call if it is missing or unusable,
2. reads every project row from the workbook (read-only; the workbook is never modified),
3. sends **one project per API request** to Claude, with the full rules file embedded in every request,
4. checks each response for the rules' mandatory elements, then caches it as soon as it arrives,
5. rebuilds `output/Generated_Claude_Code_Project_Prompts.docx` from the cache,
6. logs every outcome, and resumes where it left off if interrupted.

Each generated prompt follows the output structure in section 14 of `prompt-rules.md`: **Part 1** (GitHub repository details) and **Part 2** (the copy-paste-ready prompt with 16 numbered sections, from expert personas through final reporting). It embeds the rules' mandatory requirements: `README.md`/`context.md`/`CHANGELOG.md` maintenance, a new-session continuity protocol, Git identity and repository protection, milestone-wise commits, a free-first policy, evidence-based testing, and a security review.

## Features

- **Rules gate:** `prompt-rules.md` is read and validated before generation, sent verbatim with every request, and its SHA-256 is stored with each cached prompt. Responses missing mandatory elements (Part 1/Part 2, README/context/CHANGELOG, milestones, Git identity, the LinkedIn signal rule, the supplied repo name) are rejected and regenerated once.
- **Rule changes:** prompts made under an older rules file are reported and regenerated only with `--apply-new-rules`, so changing the rules never spends money silently.
- Intelligent column detection (`Project Name`, `project_name`, `PROJECT NAME`, `Name`, ... all work); ambiguous or missing columns stop processing with a clear message.
- Automatic worksheet selection, or `EXCEL_SHEET_NAME` / `--sheet` when several sheets qualify.
- Skips empty rows, rows without a project name, and exact duplicate rows, and logs them.
- One streamed API request per project, with configurable retries, exponential backoff, `retry-after` support, and timeouts.
- Fatal errors (invalid API key, unknown model, no permission, billing) stop the run immediately instead of failing every project.
- Truncated, empty, too-short, or refused responses are never saved as successes.
- Persistent prompt cache plus progress file with atomic writes. Completed projects are never regenerated unless their Excel data changes or you ask (`--force`).
- Change detection: editing a project's row regenerates only that project.
- The Word document is rebuilt from the cache (no duplicates, no API calls needed), with a title page, table of contents, Heading 1/2 structure, bullets, numbered lists, tables, and monospace code blocks.
- CSV processing log (`SUCCESS` / `FAILED` / `SKIPPED`) and an application log, with secrets redacted.
- Token usage reporting per project and per run.
- CLI options: `--dry-run`, `--retry-failed`, `--rebuild-document`, `--limit N`, `--force ROW ...`, `--sheet NAME`.

## Technology stack

| Technology | Purpose |
| --- | --- |
| Python 3.11+ (tested on 3.14) | Application language |
| `anthropic` (official SDK) | Claude API requests, typed errors, streaming |
| `openpyxl` | Read-only Excel access |
| `python-docx` | Word document generation |
| `python-dotenv` | `.env` configuration |
| `pytest` | Automated tests |

There are no databases, web frameworks, or paid services besides the Anthropic API itself.

## Prerequisites

- Python 3.11 or newer
- An Anthropic API key with available credits (API usage is billed; see [Cost](#cost))
- `Projects_Master_List.xlsx` (or `Project_Master_List.xlsx`) and `prompt-rules.md` in the project root

## Installation

Windows (PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

Linux / macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## API configuration

1. Create an API key at <https://console.anthropic.com/settings/keys> (sign in, then **Create Key**). Keep it private and never paste it into chats, issues, or code.
2. Copy `.env.example` to `.env` in the project root.
3. Fill in at least:

```dotenv
ANTHROPIC_API_KEY=your_api_key_here
ANTHROPIC_MODEL=claude-opus-5-5
```

`.env` is listed in `.gitignore` and must never be committed.

### Choosing a model

| Model ID | Quality | Price (input / output per 1M tokens)* |
| --- | --- | --- |
| `claude-opus-5-5` | Highest quality for detailed engineering specs (**recommended**) | $4 / $20 |
| `claude-sonnet-5-5` | Strong quality at half the price | $2 / $10 |
| `claude-haiku-5-5` | Fastest and cheapest, less thorough | $0.10 / $0.50 |

\*Anthropic first-party API list prices as of October 2026. Prices change, so check <https://www.anthropic.com/pricing> before large runs.

### All settings

| Variable | Default | Meaning |
| --- | --- | --- |
| `ANTHROPIC_API_KEY` | *(required)* | Your API key |
| `ANTHROPIC_MODEL` | *(required)* | Model ID, e.g. `claude-opus-5-5` |
| `ANTHROPIC_EFFORT` | `medium` | `low` / `medium` / `high` / `xhigh` / `max`; `default` omits the parameter (needed for older models such as Haiku 4.5) |
| `ENABLE_REFUSAL_FALLBACK` | `true` | On Opus 5.5 / Opus 5 / Sonnet 5.5 / Fable 5.1, enables Anthropic's server-side fallback if a request is declined by a safety classifier |
| `MAX_RETRIES` | `4` | Retries after the first attempt (so 5 attempts total) |
| `RETRY_BASE_DELAY` | `2` | Seconds; waits are 2, 4, 8, 16 ... (capped at 60s; `retry-after` takes precedence) |
| `API_TIMEOUT` | `600` | Seconds per request |
| `MAX_TOKENS` | `32000` | Output token ceiling per project (includes Claude's internal reasoning) |
| `EXCEL_FILE` | *(auto)* | Uses `Project_Master_List.xlsx` or `Projects_Master_List.xlsx`, whichever exists; relative paths resolve from the project root |
| `PROMPT_RULES_FILE` | `prompt-rules.md` | Mandatory rules file sent with every request |
| `EXCEL_SHEET_NAME` | *(auto)* | Worksheet to read when several contain project data |
| `OUTPUT_DIR` / `LOG_DIR` / `DATA_DIR` | `output` / `logs` / `data` | Output locations |
| `MAX_PROJECTS` | *(unlimited)* | Maximum projects per run |
| `REQUEST_DELAY` | `0` | Seconds to pause between requests |

### Cost

API usage is billed by Anthropic. Each request sends a fixed system prompt (about 2,600 tokens) plus the project's fields, and receives a long prompt (typically several thousand tokens) plus internal reasoning tokens, which are billed as output. As a rough estimate only, a project costs on the order of **$0.10 to $0.50 with Opus 5.5 at `medium` effort** (less with Sonnet 5.5). Measure with a small run first:

```bash
python main.py --limit 2
```

The run summary prints the exact token usage, and `data/progress.json` keeps running totals. The system prompt is marked for prompt caching. If it is shorter than the model's minimum cacheable length, caching simply does not apply; the summary shows how many input tokens came from the cache.

## Excel workbook

- **File:** `Projects_Master_List.xlsx` or `Project_Master_List.xlsx` in the project root (or set `EXCEL_FILE`).
- **Columns** (header names are matched case- and underscore-insensitively; the header row may be anywhere in the first 20 rows):

| Logical column | Accepted header names |
| --- | --- |
| Project Idea (optional) | Project Idea, Project Description, Idea |
| Project Name | Project Name, Name, Project Title |
| Client Problem Solved | Client Problem Solved, Problem Solved, Client Problem |
| GitHub Repo Name | GitHub Repo Name, GitHub Repository Name, Repository Name, Repo Name, GitHub Repo |
| GitHub Repo Description | GitHub Repo Description, GitHub Repository Description, Repository Description, Repo Description |

- **Extra columns** (Project ID, URL, status, notes, ...) are ignored. Without a Project Idea column, the project is described by its name, client problem, and repository description.
- **Rows:** completely empty rows are ignored; rows without a Project Name are skipped and logged; identical duplicate rows are skipped. Blank optional cells are passed to Claude as "not provided"; nothing is invented.
- **Formulas:** the cached value Excel last calculated is used. Formulas in a file never opened in Excel read as blank.
- **Worksheets:** if exactly one sheet has all columns it is used automatically; otherwise set `EXCEL_SHEET_NAME` or pass `--sheet NAME`.

## Running

```bash
python main.py                      # process all new, changed, and previously failed projects
python main.py --dry-run            # validate config + workbook; list what would run; no API calls
python main.py --limit 5            # process at most 5 projects this run
python main.py --retry-failed       # process only projects whose last attempt failed
python main.py --force 7 12         # regenerate the projects on Excel rows 7 and 12
python main.py --rebuild-document   # rebuild the Word file from cached prompts; no API calls
python main.py --apply-new-rules    # regenerate prompts made under an older prompt-rules.md
python main.py --sheet "2026 Ideas" # choose a worksheet
```

Exit codes: `0` success, `1` some projects failed, `2` configuration/workbook problem, `130` interrupted.

## Output

| Path | Contents |
| --- | --- |
| `output/Generated_Claude_Code_Project_Prompts.docx` | Title page, table of contents, one section per successful project |
| `logs/processing_log.csv` | One row per outcome: project number, Excel row, name, identifier, status, timestamp, error |
| `logs/app.log` | Diagnostic log (retries, errors, run summaries) |
| `data/generated_prompts.json` | Cache of every generated prompt (source of truth for the document) |
| `data/progress.json` | Per-project status, attempts, last error, token totals, last-run summary |

**Table of contents:** Word builds it when the document is opened. If Word asks to update fields, choose **Yes**; otherwise right-click the table and choose **Update Field** (or press F9).

## Resume and duplicate prevention

- Every project gets a stable identifier from its normalized **Project Name + GitHub Repo Name**, so inserting, deleting, or reordering rows does not cause regeneration.
- A content hash of all five fields detects edits. A changed project is regenerated; until it succeeds, its old prompt is kept in the cache but left out of the document.
- Each prompt is written to the cache (atomically: temporary file, validate, replace) **before** anything else, so an interruption never loses a paid-for result.
- The Word document is rebuilt from the cache after every success, so it never contains duplicates and can always be regenerated with `--rebuild-document`.
- Press **Ctrl+C** at any time; rerun `python main.py` to continue.

## Testing

```bash
python -m pytest -q
```

All automated tests use mocked API responses and temporary workbooks. They need no API key and make no network calls. They cover the Excel reader, the generator (retries, rate limits, auth errors, timeouts, empty/truncated/refused responses), persistence (resume, change detection, corruption recovery, atomic writes), document generation, and end-to-end workflows (crash and resume, network loss, Word failure, new and changed projects).

## Troubleshooting

| Problem | What to do |
| --- | --- |
| **Prompt rules file not found / empty** | Generation never starts without `prompt-rules.md`. Restore it in the project root or set `PROMPT_RULES_FILE`. |
| **"does not comply with prompt-rules.md"** | The response lacked a mandatory element twice. Rerun to retry the project; if it persists, raise `ANTHROPIC_EFFORT` or `MAX_TOKENS`. |
| **Missing API key** | Create `.env` from `.env.example`, set `ANTHROPIC_API_KEY`, rerun. |
| **Invalid API key** (`Authentication failed`) | The run stops immediately. Check the key in the Console and in `.env` (no quotes or spaces). |
| **Model not found** | Use an exact model ID such as `claude-opus-5-5` (not a display name like `Opus-5.5`), and check that your organization can use it. Retired IDs such as `claude-3-5-sonnet-20241022` no longer work. |
| **Missing workbook** | Put `Project_Master_List.xlsx` in the project root or set `EXCEL_FILE`. |
| **Missing / ambiguous columns** | The message names the column and worksheet. Rename headers to the accepted names above. |
| **Multiple worksheets** | Set `EXCEL_SHEET_NAME` or use `--sheet`. |
| **Rate limits (429) / overloaded (529)** | Retried automatically. For persistent limits, raise `REQUEST_DELAY` or `RETRY_BASE_DELAY`, then rerun (failed projects are retried). |
| **Network errors / timeouts** | Retried automatically; projects that still fail are retried on the next run. Raise `API_TIMEOUT` if needed. |
| **"Response was cut off at MAX_TOKENS"** | Increase `MAX_TOKENS` (e.g. `48000`) and rerun. |
| **Word document errors** | If the file is open in Word it cannot be replaced. Close it and run `python main.py --rebuild-document`. Prompts are already cached. |
| **Corrupted `progress.json`** | Handled automatically: the file is moved to `progress.json.corrupt-<timestamp>` and rebuilt from the prompt cache. |
| **Corrupted `generated_prompts.json`** | The tool stops without overwriting it. Restore it from a backup, or move it aside to start fresh (prompts would be regenerated, at a cost). |
| **Missing dependencies** (`ModuleNotFoundError`) | Activate `.venv` and run `pip install -r requirements.txt`. |

## Limitations

- Live Anthropic API generation has only been verified with mocked responses in this repository's tests. Do a small live run (`--limit 1`) first.
- The Markdown-to-Word converter handles headings, bullet and numbered lists, tables, code blocks, quotes, bold/italic/inline code, and links. Unusual Markdown may render as plain text.
- Numbered lists keep Claude's own numbers as text, not Word auto-numbering.
- `output/`, `logs/`, `data/`, and the workbook are git-ignored by default because they may contain private project information. Back up `data/generated_prompts.json` if your prompts matter to you.

## Deployment considerations

This is a local command-line tool and needs no server. To run it on another machine, copy the repository, create `.venv`, install requirements, and provide `.env` and the workbook. For scheduled runs, call `python main.py` from Task Scheduler or cron. It is safe to rerun because completed projects are skipped.
