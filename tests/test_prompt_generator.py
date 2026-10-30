"""Claude generator and prompt template tests (all mocked; no live API calls)."""

from pathlib import Path

import anthropic
import pytest

from src.claude_generator import (
    FALLBACK_BETA,
    ClaudeGenerator,
    FatalAPIError,
    GenerationError,
    extract_prompt_text,
)
from src.excel_reader import Project
from src.prompt_template import (
    BLANK_MARKER,
    RulesError,
    build_system_prompt,
    build_user_message,
    check_compliance,
    load_prompt_rules,
)
from tests.conftest import (
    SAMPLE_PROMPT,
    FakeClient,
    api_error,
    connection_error,
    fake_message,
    timeout_error,
)

PROJECT = Project(
    excel_row=2,
    project_idea="Match blood donors with hospitals",
    project_name="BloodBridge AI",
    client_problem="Hospitals cannot find donors fast enough",
    github_repo_name="bloodbridge-ai",
    github_repo_description="AI donor matching platform",
)


def make_generator(config, rules, outcomes=None):
    client = FakeClient(outcomes)
    sleeps = []
    return ClaudeGenerator(config, rules, client=client, sleep=sleeps.append), client, sleeps


def test_request_contains_all_project_information(config, rules):
    generator, client, _ = make_generator(config, rules)
    generator.generate(PROJECT)
    request = client.requests[0]
    user_text = request["messages"][0]["content"]
    for value in PROJECT.fields().values():
        assert value in user_text
    assert request["model"] == "claude-test-model"
    assert request["max_tokens"] == config.max_tokens
    assert request["system"][0]["text"] == build_system_prompt(rules)
    assert request["system"][0]["cache_control"] == {"type": "ephemeral"}
    assert request["output_config"] == {"effort": "medium"}
    assert len(client.requests) == 1  # one request per project


def test_rules_are_embedded_verbatim_in_system_prompt(rules):
    system = build_system_prompt(rules)
    assert rules.text in system
    assert "<prompt_rules>" in system and "the rules win" in system
    for heading in ["Part 1 - GitHub Repository Details", "Part 2 - Complete Copy-Paste-Ready Coding Prompt",
                    "### 1. Expert Personas", "### 14. Git Identity", "### 16. Final Reporting",
                    "DONE — GO TO LINKEDIN"]:
        assert heading in system
    assert build_system_prompt(rules) == system  # static per run -> cacheable prefix


def test_real_rules_file_loads():
    real = Path(__file__).resolve().parent.parent / "prompt-rules.md"
    if not real.exists():
        pytest.skip("prompt-rules.md not present")
    rules = load_prompt_rules(real)
    assert "Coding Prompt Output Structure" in rules.text and len(rules.sha256) == 64


@pytest.mark.parametrize(
    "content,match",
    [(None, "not found"), ("", "empty"), ("just text, no headings", "no Markdown headings")],
)
def test_unusable_rules_file_is_rejected(tmp_path, content, match):
    path = tmp_path / "prompt-rules.md"
    if content is not None:
        path.write_text(content, encoding="utf-8")
    with pytest.raises(RulesError, match=match):
        load_prompt_rules(path)


def test_compliance_check_lists_missing_mandatory_elements():
    missing = check_compliance("# Title\n\nSome prompt without the required parts.", PROJECT)
    for item in ["Part 1", "README.md", "context.md", "CHANGELOG.md", "milestone", "Git identity",
                 "GO TO LINKEDIN", "bloodbridge-ai"]:
        assert any(item in m for m in missing), item
    assert check_compliance(SAMPLE_PROMPT.replace("{repo}", "bloodbridge-ai"), PROJECT) == []


def test_non_compliant_response_retried_once_then_fails(config, rules):
    bad = fake_message(text="# Prompt\n\n" + "Generic text without mandatory parts. " * 60)
    generator, client, _ = make_generator(config, rules, [bad, bad, bad])
    with pytest.raises(GenerationError, match="does not comply with prompt-rules.md"):
        generator.generate(PROJECT)
    assert len(client.requests) == 2  # one retry, not the full MAX_RETRIES budget


def test_blank_fields_are_marked_not_invented():
    project = Project(3, "", "Name Only", "", "", "")
    message = build_user_message(project)
    assert message.count(BLANK_MARKER) == 3  # missing Project Idea is omitted, not marked
    assert "PROJECT IDEA" not in message
    assert "<project_name>Name Only</project_name>" in message


def test_field_content_is_escaped():
    project = Project(3, "uses <script> & </project_idea> tags", "X", "", "", "")
    message = build_user_message(project)
    assert "&lt;/project_idea&gt;" in message
    assert message.count("</project_idea>") == 1


def test_successful_response(config, rules):
    generator, _, _ = make_generator(config, rules)
    result = generator.generate(PROJECT)
    assert result.text.startswith("# Claude Code Prompt")
    assert result.attempts == 1
    assert result.usage["output_tokens"] == 2000
    assert result.request_id == "req_test"


def test_fallback_enabled_only_for_supported_models(config, rules):
    config.model = "claude-opus-5-5"
    generator, client, _ = make_generator(config, rules)
    generator.generate(PROJECT)
    assert client.requests[0]["betas"] == [FALLBACK_BETA]
    assert client.requests[0]["fallbacks"] == "default"

    config.model = "claude-haiku-4-5"
    config.effort = "default"
    generator, client, _ = make_generator(config, rules)
    generator.generate(PROJECT)
    assert "fallbacks" not in client.requests[0]
    assert "output_config" not in client.requests[0]


def test_rate_limit_retries_with_backoff_then_succeeds(config, rules):
    outcomes = [
        api_error(anthropic.RateLimitError, 429),
        api_error(anthropic.InternalServerError, 500),
        api_error(anthropic.APIStatusError, 529),
        fake_message(),
    ]
    generator, client, sleeps = make_generator(config, rules, outcomes)
    retries = []
    result = generator.generate(PROJECT, on_retry=lambda *a: retries.append(a))
    assert result.attempts == 4
    assert sleeps == [2.0, 4.0, 8.0]
    assert [r[0] for r in retries] == [2, 3, 4]


def test_retry_after_header_is_respected(config, rules):
    outcomes = [api_error(anthropic.RateLimitError, 429, {"retry-after": "17"}), fake_message()]
    generator, _, sleeps = make_generator(config, rules, outcomes)
    generator.generate(PROJECT)
    assert sleeps == [17.0]


def test_retry_limit_is_enforced(config, rules):
    config.max_retries = 2
    outcomes = [connection_error(), timeout_error(), connection_error(), fake_message()]
    generator, client, sleeps = make_generator(config, rules, outcomes)
    with pytest.raises(GenerationError) as exc:
        generator.generate(PROJECT)
    assert "gave up after 3 attempt(s)" in str(exc.value)
    assert len(client.requests) == 3
    assert len(sleeps) == 2


def test_timeout_is_retried(config, rules):
    generator, _, _ = make_generator(config, rules, [timeout_error(), fake_message()])
    assert generator.generate(PROJECT).attempts == 2


@pytest.mark.parametrize(
    "error_cls,status",
    [
        (anthropic.AuthenticationError, 401),
        (anthropic.PermissionDeniedError, 403),
        (anthropic.NotFoundError, 404),
        (anthropic.APIStatusError, 402),
    ],
)
def test_permanent_errors_are_fatal_and_not_retried(config, rules, error_cls, status):
    generator, client, sleeps = make_generator(config, rules, [api_error(error_cls, status)])
    with pytest.raises(FatalAPIError):
        generator.generate(PROJECT)
    assert len(client.requests) == 1
    assert sleeps == []


def test_bad_request_fails_project_without_retry(config, rules):
    generator, client, _ = make_generator(config, rules, [api_error(anthropic.BadRequestError, 400)])
    with pytest.raises(GenerationError) as exc:
        generator.generate(PROJECT)
    assert not isinstance(exc.value, FatalAPIError)
    assert len(client.requests) == 1


def test_error_messages_do_not_contain_api_key(config, rules):
    generator, _, _ = make_generator(config, rules, [api_error(anthropic.AuthenticationError, 401)])
    with pytest.raises(FatalAPIError) as exc:
        generator.generate(PROJECT)
    assert config.api_key not in str(exc.value)


def test_empty_response_is_retried_then_fails(config, rules):
    config.max_retries = 1
    generator, client, _ = make_generator(config, rules, [fake_message(text=None), fake_message(text="   ")])
    with pytest.raises(GenerationError, match="empty"):
        generator.generate(PROJECT)
    assert len(client.requests) == 2


def test_short_response_is_rejected(config, rules):
    config.max_retries = 0
    generator, _, _ = make_generator(config, rules, [fake_message(text="# Too short")])
    with pytest.raises(GenerationError, match="unexpectedly short"):
        generator.generate(PROJECT)


def test_truncated_response_is_not_accepted(config, rules):
    generator, client, _ = make_generator(config, rules, [fake_message(stop_reason="max_tokens")])
    with pytest.raises(GenerationError, match="MAX_TOKENS"):
        generator.generate(PROJECT)
    assert len(client.requests) == 1


def test_refusal_is_reported(config, rules):
    generator, _, _ = make_generator(config, rules, [fake_message(stop_reason="refusal")])
    with pytest.raises(GenerationError, match="declined"):
        generator.generate(PROJECT)


def test_preamble_before_first_heading_is_removed():
    message = fake_message(text="Here is your prompt:\n\n" + SAMPLE_PROMPT)
    assert extract_prompt_text(message).startswith("# Claude Code Prompt")
