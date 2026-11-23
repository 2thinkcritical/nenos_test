import json
from pathlib import Path

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
