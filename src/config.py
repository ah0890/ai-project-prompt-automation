"""Configuration loading and validation.

All settings come from environment variables, optionally loaded from a `.env`
file in the project root. Secrets are never printed: `Config.__repr__` hides
the API key, and validation messages only say whether a key is present.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Workbook names looked for in the project root when EXCEL_FILE is not set.
WORKBOOK_CANDIDATES = ("Project_Master_List.xlsx", "Projects_Master_List.xlsx")

PLACEHOLDER_KEYS = {"your_api_key_here", "your-api-key", "sk-ant-...", "changeme"}
# "default" omits output_config.effort (for older models that reject it).
VALID_EFFORTS = {"low", "medium", "high", "xhigh", "max", "default"}

MISSING_KEY_HELP = """\
ANTHROPIC_API_KEY is not configured.

To fix this:
  1. Create a file named .env in the project root:
       {root}
     (you can copy .env.example to .env)
  2. Add your Anthropic API key:
       ANTHROPIC_API_KEY=<your key from https://console.anthropic.com/settings/keys>
  3. Save the file.
  4. Run the application again:
       python main.py
"""

MISSING_MODEL_HELP = """\
ANTHROPIC_MODEL is not configured.

Add a model ID to your .env file, for example:
  ANTHROPIC_MODEL=claude-opus-5-5
See README.md ("Choosing a model") for cost/quality guidance.
"""


class ConfigError(Exception):
    """Raised when configuration is missing or invalid."""


@dataclass
class Config:
    api_key: str = field(default="", repr=False)
    model: str = ""
    max_retries: int = 4
    retry_base_delay: float = 2.0
    api_timeout: float = 600.0
    max_tokens: int = 32000
    effort: str = "medium"
    enable_refusal_fallback: bool = True
    excel_file: Path = PROJECT_ROOT / WORKBOOK_CANDIDATES[0]
    prompt_rules_file: Path = PROJECT_ROOT / "prompt-rules.md"
    excel_sheet_name: str | None = None
    output_dir: Path = PROJECT_ROOT / "output"
    log_dir: Path = PROJECT_ROOT / "logs"
    data_dir: Path = PROJECT_ROOT / "data"
    max_projects: int | None = None
    request_delay: float = 0.0

    @property
    def output_docx(self) -> Path:
        return self.output_dir / "Generated_Claude_Code_Project_Prompts.docx"

    @property
    def processing_log(self) -> Path:
        return self.log_dir / "processing_log.csv"

    @property
    def app_log(self) -> Path:
        return self.log_dir / "app.log"

    @property
    def progress_file(self) -> Path:
        return self.data_dir / "progress.json"

    @property
    def prompts_file(self) -> Path:
        return self.data_dir / "generated_prompts.json"

    def api_problems(self) -> list[str]:
        """Return human-readable problems that block API generation."""
        problems = []
        if not self.api_key or self.api_key.strip().lower() in PLACEHOLDER_KEYS:
            problems.append(MISSING_KEY_HELP.format(root=PROJECT_ROOT))
        if not self.model:
            problems.append(MISSING_MODEL_HELP)
        return problems

    def ensure_dirs(self) -> None:
        for directory in (self.output_dir, self.log_dir, self.data_dir):
            directory.mkdir(parents=True, exist_ok=True)


def _default_workbook() -> Path:
    for name in WORKBOOK_CANDIDATES:
        if (PROJECT_ROOT / name).exists():
            return PROJECT_ROOT / name
    return PROJECT_ROOT / WORKBOOK_CANDIDATES[0]


def _resolve_path(value: str, default: Path) -> Path:
    if not value:
        return default
    path = Path(value).expanduser()
    return path if path.is_absolute() else PROJECT_ROOT / path


def _opt_int(env: dict, name: str, default: int | None, minimum: int) -> int | None:
    raw = env.get(name, "").strip()
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError:
        raise ConfigError(f"{name} must be a whole number (got {raw!r}).") from None
    if value < minimum:
        raise ConfigError(f"{name} must be at least {minimum} (got {value}).")
    return value


def _int(env: dict, name: str, default: int, minimum: int) -> int:
    value = _opt_int(env, name, default, minimum)
    return default if value is None else value


def _float(env: dict, name: str, default: float, minimum: float) -> float:
    raw = env.get(name, "").strip()
    if not raw:
        return default
    try:
        value = float(raw)
    except ValueError:
        raise ConfigError(f"{name} must be a number (got {raw!r}).") from None
    if value < minimum:
        raise ConfigError(f"{name} must be at least {minimum} (got {value}).")
    return value


def _bool(env: dict, name: str, default: bool) -> bool:
    raw = env.get(name, "").strip().lower()
    if not raw:
        return default
    if raw in {"1", "true", "yes", "on"}:
        return True
    if raw in {"0", "false", "no", "off"}:
        return False
    raise ConfigError(f"{name} must be true or false (got {raw!r}).")


def load_config(env_file: Path | None = None, environ: dict | None = None) -> Config:
    """Build a Config from the environment.

    `environ` is injectable for tests; by default `.env` is loaded (without
    overriding variables already set in the shell) and `os.environ` is read.
    """
    if environ is None:
        load_dotenv(env_file or PROJECT_ROOT / ".env", override=False)
        environ = dict(os.environ)
    env = {k: (v or "") for k, v in environ.items()}

    effort = env.get("ANTHROPIC_EFFORT", "").strip().lower() or "medium"
    if effort not in VALID_EFFORTS:
        raise ConfigError(
            f"ANTHROPIC_EFFORT must be one of {sorted(VALID_EFFORTS)} (got {effort!r})."
        )

    defaults = Config()
    return Config(
        api_key=env.get("ANTHROPIC_API_KEY", "").strip(),
        model=env.get("ANTHROPIC_MODEL", "").strip(),
        max_retries=_int(env, "MAX_RETRIES", defaults.max_retries, 0),
        retry_base_delay=_float(env, "RETRY_BASE_DELAY", defaults.retry_base_delay, 0.0),
        api_timeout=_float(env, "API_TIMEOUT", defaults.api_timeout, 1.0),
        max_tokens=_int(env, "MAX_TOKENS", defaults.max_tokens, 1024),
        effort=effort,
        enable_refusal_fallback=_bool(
            env, "ENABLE_REFUSAL_FALLBACK", defaults.enable_refusal_fallback
        ),
        excel_file=_resolve_path(env.get("EXCEL_FILE", "").strip(), _default_workbook()),
        prompt_rules_file=_resolve_path(
            env.get("PROMPT_RULES_FILE", "").strip(), defaults.prompt_rules_file
        ),
        excel_sheet_name=env.get("EXCEL_SHEET_NAME", "").strip() or None,
        output_dir=_resolve_path(env.get("OUTPUT_DIR", "").strip(), defaults.output_dir),
        log_dir=_resolve_path(env.get("LOG_DIR", "").strip(), defaults.log_dir),
        data_dir=_resolve_path(env.get("DATA_DIR", "").strip(), defaults.data_dir),
        max_projects=_opt_int(env, "MAX_PROJECTS", None, 1),
        request_delay=_float(env, "REQUEST_DELAY", defaults.request_delay, 0.0),
    )
