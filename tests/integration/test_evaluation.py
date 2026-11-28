import json
from pathlib import Path

import pytest
from langchain_core.messages import AIMessage

from docx_form_filler import llm
from evaluation.run import evaluate

CORPUS = Path(__file__).resolve().parents[2] / "evaluation" / "corpus.json"


def test_fixture_corpus_fills_expected_documents(tmp_path):
    report = evaluate(CORPUS, tmp_path / "evaluation")
    assert report["cases"] >= 10
    assert report["passed"] == report["cases"]
    assert report["mode"] == "fixture"
    assert report["live_model_accuracy"] is None
    assert report["input_tokens"] == 0


def test_evaluator_detects_a_wrong_decision(tmp_path):
    case = json.loads(CORPUS.read_text())[0]
    case["decisions"][0]["value"] = "WRONG VALUE"
    corpus = tmp_path / "wrong.json"
    corpus.write_text(json.dumps([case]))
    result = evaluate(corpus, tmp_path / "result")
    assert result["passed"] == 0
    assert result["document_exact_match_rate"] == 0
    assert result["results"][0]["differences"]


@pytest.mark.parametrize("stage", ["model_initialization", "model_request"])
def test_live_provider_failure_leaves_case_unscored(tmp_path, monkeypatch, stage):
    case = json.loads(CORPUS.read_text())[0]
    corpus = tmp_path / "one.json"
    corpus.write_text(json.dumps([case]))
    output = tmp_path / "result"
    after = output / case["id"] / "after.docx"
    after.parent.mkdir(parents=True)
    after.write_bytes(b"stale artifact")

    class FailingModel:
        def __init__(self, **kwargs):
            if stage == "model_initialization":
                raise ValueError("sensitive model setup details")

        def invoke(self, messages):
            raise ConnectionError("sensitive transport details")

    monkeypatch.setattr(llm, "ChatOpenAI", FailingModel)
    result = evaluate(corpus, output, mode="live")

    assert result["live_model_accuracy"] is None
    assert result["scored_cases"] == 0
    assert result["unscored_cases"] == 1
    assert result["api_calls"] == (1 if stage == "model_request" else 0)
    assert result["successful_model_responses"] == 0
    assert result["input_tokens"] is None
    assert result["output_tokens"] is None
    assert result["usage_complete"] is False
    assert result["results"][0]["failure_stage"] == stage
    assert "sensitive" not in (output / "results.json").read_text()
    assert not after.exists()


@pytest.mark.parametrize("include_usage", [False, True])
def test_live_content_mismatch_is_scored_with_observed_usage(tmp_path, monkeypatch, include_usage):
    case = json.loads(CORPUS.read_text())[0]
    corpus = tmp_path / "one.json"
    corpus.write_text(json.dumps([case]))

    class Model:
        def __init__(self, **kwargs):
            pass

        def invoke(self, messages):
            usage = {"input_tokens": 10, "output_tokens": 2, "total_tokens": 12} if include_usage else None
            return AIMessage(content="[]", usage_metadata=usage)

    monkeypatch.setattr(llm, "ChatOpenAI", Model)
    result = evaluate(corpus, tmp_path / "result", mode="live")

    assert result["api_calls"] == result["successful_model_responses"] == 1
    assert result["scored_cases"] == 1
    assert result["unscored_cases"] == 0
    assert result["live_model_accuracy"] == 0
    assert result["input_tokens"] == (10 if include_usage else None)
    assert result["output_tokens"] == (2 if include_usage else None)
    assert result["usage_complete"] is include_usage
    assert result["results"][0]["differences"][0]["part"] == "paragraphs"


@pytest.mark.parametrize(
    "invalid", ["traversal", "absolute", "duplicate", "empty-expected", "empty-parts", "empty-label", "wrong-type"]
)
def test_invalid_later_case_is_rejected_before_any_output_writes(tmp_path, invalid):
    first, second = json.loads(CORPUS.read_text())[:2]
    if invalid in {"traversal", "absolute", "duplicate"}:
        second["id"] = {"traversal": "../escape", "absolute": str(tmp_path / "escape"), "duplicate": first["id"]}[
            invalid
        ]
    else:
        second["expected"] = {
            "empty-expected": {},
            "empty-parts": {"paragraphs": [], "tables": []},
            "empty-label": {"": ["expected text"]},
            "wrong-type": {"paragraphs": "one string"},
        }[invalid]
    corpus = tmp_path / "invalid.json"
    corpus.write_text(json.dumps([first, second]))
    output = tmp_path / "result"

    with pytest.raises(ValueError, match="case"):
        evaluate(corpus, output)

    assert not output.exists()
    assert not (tmp_path / "escape").exists()


@pytest.mark.parametrize("artifact", ["case directory", "before.docx", "input.json", "after.docx", "results.json"])
def test_evaluation_rejects_symlink_destinations_before_writing(tmp_path, artifact):
    first, second = json.loads(CORPUS.read_text())[:2]
    corpus = tmp_path / "symlink.json"
    corpus.write_text(json.dumps([first, second]))
    output = tmp_path / "result"
    output.mkdir()
    protected = tmp_path / "outside"
    protected.mkdir()
    original = protected / "unchanged"
    original.write_text("keep me")
    second_dir = output / second["id"]
    if artifact == "case directory":
        second_dir.symlink_to(protected, target_is_directory=True)
    elif artifact == "results.json":
        (output / artifact).symlink_to(original)
    else:
        second_dir.mkdir()
        (second_dir / artifact).symlink_to(original)

    with pytest.raises(ValueError, match="output"):
        evaluate(corpus, output)

    assert not (output / first["id"]).exists()
    assert list(protected.iterdir()) == [original]
    assert original.read_text() == "keep me"
