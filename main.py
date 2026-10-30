"""Claude Code Prompt Generator - command-line entry point.

    python main.py                      process all pending/changed/failed projects
    python main.py --dry-run            validate config + workbook, no API calls
    python main.py --retry-failed       only retry projects that previously failed
    python main.py --rebuild-document   rebuild the Word file from cache, no API calls
    python main.py --limit 5            process at most 5 projects this run
    python main.py --force 7 12         regenerate the projects on Excel rows 7 and 12
    python main.py --apply-new-rules    regenerate prompts made under an older prompt-rules.md
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

from src.claude_generator import ClaudeGenerator, FatalAPIError, GenerationError
from src.config import PROJECT_ROOT, Config, ConfigError, load_config
from src.document_generator import DocumentEntry, DocumentError, save_document
from src.excel_reader import WorkbookError, read_projects
from src.logger import ProcessingLog, redact, setup_logging
from src.progress_manager import (
    STATUS_FAILED,
    ProgressError,
    ProgressManager,
    TrackedProject,
    track_projects,
)
from src.prompt_template import TEMPLATE_VERSION, PromptRules, RulesError, load_prompt_rules

RULE = "=" * 50
THIN = "-" * 50

EXIT_OK, EXIT_FAILURES, EXIT_CONFIG, EXIT_INTERRUPTED = 0, 1, 2, 130


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Claude Code development prompts from Project_Master_List.xlsx."
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true",
                      help="validate configuration and workbook without making API requests")
    mode.add_argument("--rebuild-document", action="store_true",
                      help="rebuild the Word document from cached prompts (no API requests)")
    parser.add_argument("--retry-failed", action="store_true",
                        help="process only projects whose last attempt failed")
    parser.add_argument("--limit", type=int, metavar="N",
                        help="process at most N projects this run (overrides MAX_PROJECTS)")
    parser.add_argument("--force", type=int, nargs="+", metavar="ROW", default=[],
                        help="regenerate the projects on these Excel row numbers even if cached")
    parser.add_argument("--apply-new-rules", action="store_true",
                        help="also regenerate prompts that were generated under a different prompt-rules.md")
    parser.add_argument("--sheet", metavar="NAME", help="worksheet to read (overrides EXCEL_SHEET_NAME)")
    args = parser.parse_args(argv)
    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be at least 1")
    return args


def rel(path: Path) -> str:
    try:
        return path.resolve().relative_to(PROJECT_ROOT).as_posix()
    except ValueError:
        return str(path)


def document_entries(pm: ProgressManager, tracked: list[TrackedProject]) -> list[DocumentEntry]:
    entries = []
    for tp, cached in pm.valid_entries(tracked):
        fields = tp.project.fields()
        entries.append(DocumentEntry(number=tp.number, prompt=cached["prompt"], **fields))
    return entries


class Runner:
    def __init__(self, config: Config, args: argparse.Namespace, generator_factory=None, out=print):
        self.config = config
        self.args = args
        self.generator_factory = generator_factory or (lambda cfg, rules: ClaudeGenerator(cfg, rules))
        self.out = out
        self.log = logging.getLogger("prompt_generator")
        self.csv: ProcessingLog  # created in run() once the log directory exists

    def rebuild_document(self, pm, tracked, source_name: str) -> str | None:
        """Rebuild the Word document; return an error message instead of raising."""
        try:
            save_document(document_entries(pm, tracked), self.config.output_docx, source_name)
            return None
        except DocumentError as exc:
            message = redact(str(exc), (self.config.api_key,))
            self.log.error("Document generation failed: %s", message)
            return message

    def run(self) -> int:
        cfg, args, out = self.config, self.args, self.out
        out(RULE)
        out("CLAUDE CODE PROMPT GENERATOR")
        out(RULE)
        out("")

        cfg.ensure_dirs()
        secrets = (cfg.api_key,)
        self.log = setup_logging(cfg.app_log, secrets)
        self.log.info("Run started: %s", vars(args))

        # ---------------------------------------------------- configuration
        out("Checking configuration...")
        needs_api = not args.rebuild_document
        problems = cfg.api_problems() if needs_api else []
        if problems and not args.dry_run:
            out("")
            out("API configuration: INVALID")
            for problem in problems:
                out("")
                out(problem.rstrip())
            return EXIT_CONFIG
        if needs_api:
            out(f"API configuration: {'MISSING (see below)' if problems else 'VALID'}")
            if not problems:
                out(f"Model: {cfg.model} (effort: {cfg.effort}, max tokens: {cfg.max_tokens})")
        out("")

        # ------------------------------------------------------ prompt rules
        # prompt-rules.md is mandatory: it is checked before any API request.
        rules: PromptRules | None = None
        if needs_api:
            out("Checking prompt rules...")
            try:
                rules = load_prompt_rules(cfg.prompt_rules_file)
            except RulesError as exc:
                out("")
                out(f"ERROR: {exc}")
                self.log.error("Prompt rules error: %s", exc)
                return EXIT_CONFIG
            out(f"Prompt rules: {rel(rules.path)} ({rules.line_count} lines, "
                f"sha256 {rules.short_hash}) - LOADED, applied to every request")
            self.log.info("Prompt rules loaded: %s sha256=%s", rules.path.name, rules.sha256)
            out("")

        # --------------------------------------------------------- workbook
        out(f"Excel workbook: {rel(cfg.excel_file)}")
        try:
            workbook = read_projects(cfg.excel_file, args.sheet or cfg.excel_sheet_name)
        except WorkbookError as exc:
            out("")
            out(f"ERROR: {exc}")
            self.log.error("Workbook error: %s", exc)
            return EXIT_CONFIG
        out(f"Worksheet: {workbook.sheet_name} (header row {workbook.header_row})")
        out("")
        out("Loading projects...")

        tracked, duplicates = track_projects(workbook.projects)
        skipped = sorted(workbook.skipped + duplicates, key=lambda s: s.excel_row)

        pm = ProgressManager(cfg.progress_file, cfg.prompts_file)
        try:
            pm.load()
        except ProgressError as exc:
            out(f"ERROR: {exc}")
            self.log.error("Progress error: %s", exc)
            return EXIT_CONFIG
        for warning in pm.warnings:
            out(f"WARNING: {warning}")
            self.log.warning(warning)

        classes = {tp.project_id: pm.classify(tp) for tp in tracked}
        previously_completed = sum(1 for c in classes.values() if c == "completed")
        source_name = cfg.excel_file.name

        out("")
        out(f"Total projects found: {len(tracked)}")
        if skipped:
            out(f"Skipped rows: {len(skipped)}")
            for row in skipped:
                out(f"  - Excel row {row.excel_row}: {row.reason}")
        if workbook.empty_rows:
            self.log.info("Ignored %d completely empty row(s)", workbook.empty_rows)
        out(f"Previously completed: {previously_completed}")
        older_rules = (
            {tp.project_id for tp in tracked if pm.generated_under_other_rules(tp, rules.sha256)}
            if rules else set()
        )
        if older_rules:
            note = "will be regenerated" if args.apply_new_rules else "rerun with --apply-new-rules to regenerate them"
            out(f"Generated under a different prompt-rules.md: {len(older_rules)} ({note})")

        # --------------------------------------------------- rebuild only
        if args.rebuild_document:
            out("")
            out("Rebuilding Word document from cached prompts (no API requests)...")
            error = self.rebuild_document(pm, tracked, source_name)
            if error:
                out(f"FAILED: {error}")
                return EXIT_FAILURES
            out(f"Document saved with {previously_completed} project(s): {rel(cfg.output_docx)}")
            return EXIT_OK

        # ------------------------------------------------------- selection
        forced_rows = set(args.force)
        known_rows = {tp.project.excel_row for tp in tracked}
        for row in sorted(forced_rows - known_rows):
            out(f"WARNING: --force row {row} is not an eligible project row; ignored.")

        def selected(tp: TrackedProject) -> bool:
            if tp.project.excel_row in forced_rows:
                return True
            if args.apply_new_rules and tp.project_id in older_rules:
                return True
            if args.retry_failed:
                state = pm.progress["projects"].get(tp.project_id, {})
                return state.get("status") == STATUS_FAILED and state.get("content_hash") == tp.content_hash
            return classes[tp.project_id] != "completed"

        candidates = [tp for tp in tracked if selected(tp)]
        limit = args.limit or cfg.max_projects
        to_process = candidates[:limit] if limit else candidates
        out(f"Remaining projects: {len(tracked) - previously_completed}")
        if args.retry_failed:
            out(f"Previously failed (selected for retry): {len(candidates)}")
        if limit and len(candidates) > limit:
            out(f"This run is limited to {limit} project(s).")

        # --------------------------------------------------------- dry run
        if args.dry_run:
            out("")
            out(f"DRY RUN - no API requests made. Would process {len(to_process)} project(s):")
            for tp in to_process:
                reason = classes[tp.project_id]
                if reason == "completed":
                    reason = "forced" if tp.project.excel_row in forced_rows else "new rules"
                out(f"  - Project {tp.number:03d} (Excel row {tp.project.excel_row}): {tp.name} [{reason}]")
            if problems:
                out("")
                out("API configuration problems must be fixed before a real run:")
                for problem in problems:
                    out("")
                    out(problem.rstrip())
                return EXIT_CONFIG
            return EXIT_OK

        self.csv = ProcessingLog(cfg.processing_log, secrets)
        for row in skipped:
            self.csv.write("SKIPPED", excel_row=row.excel_row, project_name=row.project_name, error=row.reason)

        # ------------------------------------------------------ processing
        newly_generated = failed_this_run = 0
        usage_this_run: dict[str, int] = {}
        interrupted = fatal = False
        document_error: str | None = None

        if to_process:
            assert rules is not None  # loaded above whenever the API is needed
            generator = self.generator_factory(cfg, rules)
            out(THIN)
            for index, tp in enumerate(to_process):
                out("")
                out(f"Processing project {tp.number} of {len(tracked)}: {tp.name}")
                status = classes[tp.project_id]
                try:
                    if status == "outdated":
                        out("Source information changed since the last generation; regenerating.")
                        pm.mark_outdated(tp)
                    elif status == "completed" and tp.project.excel_row in forced_rows:
                        out("Forced regeneration requested.")
                    elif status == "completed":
                        out("Regenerating under the current prompt-rules.md.")
                    out("Generating Claude Code development prompt...")

                    def on_retry(attempt, wait, reason):
                        out(f"  Retry {attempt}/{cfg.max_retries + 1} in {wait:.0f}s ({redact(reason, secrets)})")
                        self.log.warning("Project %s retry %d: %s", tp.project_id, attempt, reason)

                    result = generator.generate(tp.project, on_retry=on_retry)
                except KeyboardInterrupt:
                    interrupted = True
                    out("")
                    out("Interrupted by user. Progress for completed projects has been saved.")
                    self.log.warning("Interrupted during project %s", tp.project_id)
                    break
                except FatalAPIError as exc:
                    message = redact(str(exc), secrets)
                    self._failure(pm, tp, message)
                    failed_this_run += 1
                    fatal = True
                    out("")
                    out("Stopping: this error affects every project, so remaining projects were not attempted.")
                    break
                except GenerationError as exc:
                    self._failure(pm, tp, redact(str(exc), secrets))
                    failed_this_run += 1
                    continue
                except Exception as exc:  # keep going on unexpected per-project errors
                    self.log.exception("Unexpected error for project %s", tp.project_id)
                    self._failure(pm, tp, redact(f"Unexpected error: {type(exc).__name__}: {exc}", secrets))
                    failed_this_run += 1
                    continue

                pm.record_success(tp, result, TEMPLATE_VERSION, rules.sha256)
                newly_generated += 1
                for key, value in result.usage.items():
                    usage_this_run[key] = usage_this_run.get(key, 0) + value
                self.csv.write("SUCCESS", tp.number, tp.project.excel_row, tp.name, tp.project_id)
                tokens = result.usage
                out(f"SUCCESS ({tokens.get('input_tokens', 0) + tokens.get('cache_read_input_tokens', 0) + tokens.get('cache_creation_input_tokens', 0)} input / "
                    f"{tokens.get('output_tokens', 0)} output tokens)")
                self.log.info("Project %s generated (%s)", tp.project_id, tokens)

                document_error = self.rebuild_document(pm, tracked, source_name)
                if document_error:
                    out(f"WARNING: prompt saved to cache, but the Word document was not updated: {document_error}")
                    self.csv.write("FAILED", tp.number, tp.project.excel_row, tp.name, tp.project_id,
                                   f"[DOCUMENT] Prompt cached; Word update failed: {document_error}")

                if cfg.request_delay and index < len(to_process) - 1:
                    try:
                        time.sleep(cfg.request_delay)
                    except KeyboardInterrupt:
                        interrupted = True
                        out("Interrupted by user. Progress has been saved.")
                        break
            out("")
            out(THIN)
        else:
            out("")
            out("Nothing to process: all eligible projects are already completed.")

        # Make sure the document reflects the cache even if nothing new succeeded
        # (e.g. first run after deleting the .docx, or outdated entries removed).
        if newly_generated == 0 or document_error:
            document_error = self.rebuild_document(pm, tracked, source_name)

        return self._summary(pm, tracked, skipped, previously_completed, newly_generated,
                             failed_this_run, usage_this_run, interrupted, fatal, document_error)

    def _failure(self, pm: ProgressManager, tp: TrackedProject, message: str) -> None:
        pm.record_failure(tp, message)
        self.csv.write("FAILED", tp.number, tp.project.excel_row, tp.name, tp.project_id, message)
        self.out(f"FAILED: {message}")
        self.log.error("Project %s failed: %s", tp.project_id, message)

    def _summary(self, pm, tracked, skipped, previously_completed, newly_generated,
                 failed_this_run, usage, interrupted, fatal, document_error) -> int:
        cfg, out = self.config, self.out
        successful = len(pm.valid_entries(tracked))
        failed = sum(
            1 for tp in tracked
            if pm.cached_entry(tp) is None
            and pm.progress["projects"].get(tp.project_id, {}).get("status") == STATUS_FAILED
            and pm.progress["projects"][tp.project_id].get("content_hash") == tp.content_hash
        )
        not_attempted = len(tracked) - successful - failed

        out("")
        out("Final summary:")
        out("")
        out(f"Total Projects: {len(tracked) + len(skipped)}")
        out(f"Successfully Generated: {successful}")
        out(f"Failed: {failed}")
        out(f"Skipped: {len(skipped)}")
        if not_attempted:
            out(f"Not Yet Processed: {not_attempted}")
        out(f"Previously Completed: {previously_completed}")
        out(f"Newly Processed: {newly_generated}")
        if failed_this_run:
            out(f"Failed This Run: {failed_this_run}")
        if usage:
            total_in = sum(usage.get(k, 0) for k in ("input_tokens", "cache_creation_input_tokens", "cache_read_input_tokens"))
            out(f"Token Usage This Run: {total_in} input "
                f"({usage.get('cache_read_input_tokens', 0)} from cache) / {usage.get('output_tokens', 0)} output")
        out("")
        if document_error:
            out(f"Output: NOT UPDATED - {document_error}")
        else:
            out("Output:")
            out(rel(cfg.output_docx))
        out("")
        out("Log:")
        out(rel(cfg.processing_log))
        out("")
        out("Progress:")
        out(rel(cfg.progress_file))
        out("")
        out(RULE)

        pm.record_run({
            "total_projects": len(tracked) + len(skipped),
            "successful": successful,
            "failed": failed,
            "skipped": len(skipped),
            "not_yet_processed": not_attempted,
            "previously_completed": previously_completed,
            "newly_generated": newly_generated,
            "failed_this_run": failed_this_run,
            "usage": usage,
            "interrupted": interrupted,
            "stopped_on_fatal_error": fatal,
            "document_error": document_error,
        })
        self.log.info("Run finished: %d successful, %d failed, %d new", successful, failed, newly_generated)

        if interrupted:
            return EXIT_INTERRUPTED
        if fatal or failed_this_run or document_error:
            return EXIT_FAILURES
        return EXIT_OK


def main(argv: list[str] | None = None, environ: dict | None = None, generator_factory=None, out=print) -> int:
    args = parse_args(argv)
    try:
        config = load_config(environ=environ)
    except ConfigError as exc:
        out(f"Configuration error: {exc}")
        return EXIT_CONFIG
    try:
        return Runner(config, args, generator_factory, out).run()
    except KeyboardInterrupt:
        out("\nInterrupted. Completed projects are saved; rerun to continue.")
        return EXIT_INTERRUPTED


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    sys.exit(main())
