"""Anthropic Claude API integration: one request per project.

Retries are handled here (the SDK's own retries are disabled) so that
MAX_RETRIES / RETRY_BASE_DELAY are honoured exactly:

    attempt 1: immediate; attempt 2: wait base; attempt 3: wait 2*base; ...

A server-provided `retry-after` header takes precedence over the computed delay.
Authentication, permission, and unknown-model errors are fatal for the whole
run (retrying them, or moving on to the next project, cannot succeed).

Every request is made under the loaded prompt-rules.md, and every response is
checked for the rules' mandatory elements before it is accepted.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Callable

import anthropic

from src.config import Config
from src.excel_reader import Project
from src.prompt_template import (
    PromptRules,
    build_system_prompt,
    build_user_message,
    check_compliance,
)

MAX_BACKOFF_SECONDS = 60.0
MAX_RETRY_AFTER_SECONDS = 300.0
MIN_PROMPT_CHARS = 1500
# A non-compliant response is regenerated at most this many times (each costs a full request).
MAX_COMPLIANCE_RETRIES = 1
RETRYABLE_STATUS = {408, 409}

# Models that accept the server-side refusal fallback in its "default" form.
FALLBACK_MODELS = {"claude-fable-5-1", "claude-opus-5-5", "claude-opus-5", "claude-sonnet-5-5"}
FALLBACK_BETA = "server-side-fallback-2026-07-01"


class GenerationError(Exception):
    """A single project's generation failed; processing can continue."""

    def __init__(self, message: str, retryable: bool = False):
        super().__init__(message)
        self.retryable = retryable


class FatalAPIError(GenerationError):
    """A configuration-level failure (bad key, model, permissions): stop the run."""


@dataclass
class GenerationResult:
    text: str
    model: str
    stop_reason: str | None
    attempts: int
    request_id: str | None = None
    usage: dict[str, int] = field(default_factory=dict)


def _usage_dict(message) -> dict[str, int]:
    usage = getattr(message, "usage", None)
    keys = ("input_tokens", "output_tokens", "cache_creation_input_tokens", "cache_read_input_tokens")
    return {k: int(getattr(usage, k, 0) or 0) for k in keys} if usage else {}


def _retry_after_seconds(error: anthropic.APIStatusError) -> float | None:
    try:
        raw = error.response.headers.get("retry-after")
        return min(float(raw), MAX_RETRY_AFTER_SECONDS) if raw is not None else None
    except (AttributeError, TypeError, ValueError):
        return None


def _status_detail(error: anthropic.APIStatusError) -> str:
    request_id = getattr(error, "request_id", None)
    suffix = f", request id {request_id}" if request_id else ""
    return f"HTTP {error.status_code}{suffix}: {error.message}"


def extract_prompt_text(message) -> str:
    """Join text blocks and drop any chatter before the first Markdown heading."""
    parts = [b.text for b in (getattr(message, "content", None) or []) if getattr(b, "type", "") == "text"]
    text = "\n".join(p for p in parts if p).strip()
    if text and not text.startswith("#"):
        heading = text.find("\n#")
        if 0 <= heading < 600:
            text = text[heading + 1 :].strip()
    return text


class ClaudeGenerator:
    def __init__(
        self,
        config: Config,
        rules: PromptRules,
        client: anthropic.Anthropic | None = None,
        sleep: Callable[[float], None] = time.sleep,
    ):
        self.config = config
        self.client = client or anthropic.Anthropic(
            api_key=config.api_key, timeout=config.api_timeout, max_retries=0
        )
        self.sleep = sleep
        self.rules = rules
        self.system_prompt = build_system_prompt(rules)
        self._compliance_failures = 0
        self.use_fallback = config.enable_refusal_fallback and config.model in FALLBACK_MODELS

    def build_request(self, project: Project) -> dict:
        request = {
            "model": self.config.model,
            "max_tokens": self.config.max_tokens,
            "system": [
                {"type": "text", "text": self.system_prompt, "cache_control": {"type": "ephemeral"}}
            ],
            "messages": [{"role": "user", "content": build_user_message(project)}],
        }
        if self.config.effort != "default":
            request["output_config"] = {"effort": self.config.effort}
        if self.use_fallback:
            request["betas"] = [FALLBACK_BETA]
            request["fallbacks"] = "default"
        return request

    def _call(self, request: dict):
        api = self.client.beta.messages if "betas" in request else self.client.messages
        # Streaming avoids HTTP timeouts on long outputs with large max_tokens.
        with api.stream(**request) as stream:
            return stream.get_final_message()

    def _backoff(self, attempt: int) -> float:
        return min(self.config.retry_base_delay * (2 ** (attempt - 1)), MAX_BACKOFF_SECONDS)

    def generate(
        self, project: Project, on_retry: Callable[[int, float, str], None] | None = None
    ) -> GenerationResult:
        request = self.build_request(project)
        self._compliance_failures = 0
        max_attempts = self.config.max_retries + 1
        last_error = "unknown error"

        for attempt in range(1, max_attempts + 1):
            delay: float | None = None
            try:
                message = self._call(request)
                return self._validate(message, attempt, project)
            except anthropic.AuthenticationError as e:
                raise FatalAPIError(
                    f"Authentication failed ({_status_detail(e)}). The API key in "
                    "ANTHROPIC_API_KEY was rejected; check it in .env."
                ) from None
            except anthropic.PermissionDeniedError as e:
                raise FatalAPIError(
                    f"Permission denied ({_status_detail(e)}). Check your Anthropic "
                    "organization/workspace access for this model."
                ) from None
            except anthropic.NotFoundError as e:
                raise FatalAPIError(
                    f"Model or endpoint not found ({_status_detail(e)}). Check "
                    f"ANTHROPIC_MODEL (currently {self.config.model!r})."
                ) from None
            except anthropic.RateLimitError as e:
                last_error = f"Rate limited ({_status_detail(e)})"
                delay = _retry_after_seconds(e)
            except anthropic.APIStatusError as e:
                last_error = f"API error ({_status_detail(e)})"
                if e.status_code == 402:
                    raise FatalAPIError(
                        f"Billing problem ({_status_detail(e)}). Check your Anthropic account credits."
                    ) from None
                if e.status_code < 500 and e.status_code not in RETRYABLE_STATUS:
                    raise GenerationError(last_error) from None
                delay = _retry_after_seconds(e)
            except anthropic.APITimeoutError:
                last_error = f"Request timed out after {self.config.api_timeout:g}s"
            except anthropic.APIConnectionError as e:
                last_error = f"Network/connection error: {e}"
            except GenerationError as e:
                if not e.retryable:
                    raise
                last_error = str(e)

            if attempt < max_attempts:
                wait = delay if delay is not None else self._backoff(attempt)
                if on_retry:
                    on_retry(attempt + 1, wait, last_error)
                self.sleep(wait)

        raise GenerationError(f"{last_error} (gave up after {max_attempts} attempt(s))")

    def _validate(self, message, attempt: int, project: Project) -> GenerationResult:
        stop_reason = getattr(message, "stop_reason", None)
        if stop_reason == "refusal":
            details = getattr(message, "stop_details", None)
            category = getattr(details, "category", None) if details else None
            raise GenerationError(
                f"Claude declined to generate this prompt (refusal, category: {category or 'unspecified'})."
            )
        if stop_reason == "max_tokens":
            raise GenerationError(
                f"Response was cut off at MAX_TOKENS={self.config.max_tokens}; "
                "increase MAX_TOKENS in .env and retry this project."
            )
        text = extract_prompt_text(message)
        if not text:
            raise GenerationError("Claude returned an empty response.", retryable=True)
        if len(text) < MIN_PROMPT_CHARS:
            raise GenerationError(
                f"Claude returned an unexpectedly short response ({len(text)} characters).",
                retryable=True,
            )
        missing = check_compliance(text, project)
        if missing:
            self._compliance_failures += 1
            raise GenerationError(
                "Generated prompt does not comply with prompt-rules.md; missing: "
                + ", ".join(missing),
                retryable=self._compliance_failures <= MAX_COMPLIANCE_RETRIES,
            )
        return GenerationResult(
            text=text,
            model=getattr(message, "model", None) or self.config.model,
            stop_reason=stop_reason,
            attempts=attempt,
            request_id=getattr(message, "_request_id", None),
            usage=_usage_dict(message),
        )
