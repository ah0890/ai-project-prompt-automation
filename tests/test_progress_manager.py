import json
from dataclasses import replace

import pytest

from src.claude_generator import GenerationResult
from src.excel_reader import Project
from src.progress_manager import (
    STATUS_FAILED,
    ProgressError,
    ProgressManager,
    atomic_write_json,
    track_projects,
)


def projects(n=3):
    return [
        Project(i + 2, f"idea {i}", f"Project {i}", f"problem {i}", f"repo-{i}", f"desc {i}")
        for i in range(n)
    ]


def result(text="# Prompt\n\nbody"):
    return GenerationResult(text=text, model="m", stop_reason="end_turn", attempts=1,
                            usage={"input_tokens": 10, "output_tokens": 20})


@pytest.fixture
def pm(tmp_path):
    manager = ProgressManager(tmp_path / "data" / "progress.json", tmp_path / "data" / "generated_prompts.json")
    manager.load()
    return manager


def reload(pm):
    fresh = ProgressManager(pm.progress_path, pm.prompts_path)
    fresh.load()
    return fresh


def test_save_and_load_progress(pm):
    tracked, _ = track_projects(projects())
    pm.record_success(tracked[0], result(), "1.0", "rules-hash")
    pm.record_failure(tracked[1], "boom")
    loaded = reload(pm)
    assert loaded.classify(tracked[0]) == "completed"
    assert loaded.classify(tracked[1]) == "failed"
    assert loaded.classify(tracked[2]) == "new"
    assert loaded.progress["projects"][tracked[1].project_id]["last_error"] == "boom"
    assert loaded.progress["usage_totals"]["output_tokens"] == 20
    assert loaded.cached_entry(tracked[0])["prompt"] == "# Prompt\n\nbody"


def test_ids_are_stable_when_rows_move():
    original, _ = track_projects(projects())
    moved = [replace(p, excel_row=p.excel_row + 10) for p in reversed(projects())]
    shifted, _ = track_projects(moved)
    assert {t.project_id for t in original} == {t.project_id for t in shifted}
    assert {t.content_hash for t in original} == {t.content_hash for t in shifted}


def test_changed_project_is_detected_as_outdated(pm):
    tracked, _ = track_projects(projects())
    pm.record_success(tracked[0], result(), "1.0", "rules-hash")
    changed = [replace(projects()[0], client_problem="a brand new problem")] + projects()[1:]
    retracked, _ = track_projects(changed)
    assert retracked[0].project_id == tracked[0].project_id
    assert pm.classify(retracked[0]) == "outdated"
    assert pm.cached_entry(retracked[0]) is None  # outdated prompt excluded from document
    pm.mark_outdated(retracked[0])
    assert reload(pm).prompts[tracked[0].project_id]["outdated"] is True


def test_whitespace_only_changes_do_not_trigger_regeneration():
    a, _ = track_projects(projects(1))
    b, _ = track_projects([replace(projects(1)[0], project_idea="idea   0 ")])
    assert a[0].content_hash == b[0].content_hash


def test_failed_project_can_succeed_later(pm):
    tracked, _ = track_projects(projects(1))
    pm.record_failure(tracked[0], "network down")
    pm.record_failure(tracked[0], "network down again")
    assert pm.progress["projects"][tracked[0].project_id]["attempts"] == 2
    pm.record_success(tracked[0], result(), "1.0", "rules-hash")
    loaded = reload(pm)
    assert loaded.classify(tracked[0]) == "completed"
    assert loaded.progress["projects"][tracked[0].project_id]["status"] == "SUCCESS"


def test_corrupt_progress_is_backed_up_and_rebuilt_from_cache(pm):
    tracked, _ = track_projects(projects())
    pm.record_success(tracked[0], result(), "1.0", "rules-hash")
    pm.progress_path.write_text("{ not json", encoding="utf-8")
    loaded = reload(pm)
    assert loaded.warnings and "rebuilt" in loaded.warnings[0]
    assert loaded.classify(tracked[0]) == "completed"
    assert loaded.progress["projects"][tracked[0].project_id]["status"] == "SUCCESS"
    assert list(pm.progress_path.parent.glob("progress.json.corrupt-*"))


def test_corrupt_prompt_cache_is_never_overwritten(pm):
    pm.prompts_path.parent.mkdir(parents=True, exist_ok=True)
    pm.prompts_path.write_text("{ truncated", encoding="utf-8")
    with pytest.raises(ProgressError, match="NOT overwritten"):
        reload(pm)
    assert pm.prompts_path.read_text(encoding="utf-8") == "{ truncated"


def test_atomic_write_keeps_previous_file_when_write_fails(tmp_path, monkeypatch):
    target = tmp_path / "state.json"
    atomic_write_json(target, {"good": True})

    def explode(*_args, **_kwargs):
        raise OSError("disk full")

    monkeypatch.setattr("src.progress_manager.os.replace", explode)
    with pytest.raises(OSError):
        atomic_write_json(target, {"good": False})
    assert json.loads(target.read_text(encoding="utf-8")) == {"good": True}
    assert [p.name for p in tmp_path.iterdir()] == ["state.json"]  # temp file cleaned up


def test_failed_state_reset_when_project_changes(pm):
    tracked, _ = track_projects(projects(1))
    pm.record_failure(tracked[0], "boom")
    changed, _ = track_projects([replace(projects(1)[0], project_idea="new idea")])
    assert pm.classify(changed[0]) == "new"
    assert pm.progress["projects"][tracked[0].project_id]["status"] == STATUS_FAILED
