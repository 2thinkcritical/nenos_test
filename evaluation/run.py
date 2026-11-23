"""Evaluate saved DOCX contents against independent expected outputs.

Fixture mode checks application behavior with authored model responses. It does
not measure model accuracy. Live mode is opt-in and makes billable API requests.
"""

import argparse
import contextlib
import io
import json
import os
import time
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from docx import Document
from docx.oxml import OxmlElement
from dotenv import load_dotenv

import field_mapper
from docx_analyzer import get_all_placeholders, placeholder_id
from main import run_pipeline


def build_template(case: dict, target: Path) -> None:
    doc = Document()
    doc.core_properties.author = "Synthetic evaluation fixture"
    doc.core_properties.last_modified_by = "Synthetic evaluation fixture"
    for index, text in enumerate(case["paragraphs"]):
        paragraph = doc.add_paragraph()
        if index == 0:
            paragraph.style = "Title"
        if index == 1 and case.get("hyperlink"):
            hyperlink = OxmlElement("w:hyperlink")
            run = OxmlElement("w:r")
            node = OxmlElement("w:t")
            node.text = case["hyperlink"]
            run.append(node)
            hyperlink.append(run)
            paragraph._p.append(hyperlink)
        paragraph.add_run(text)
    if "header" in case:
        doc.sections[0].header.paragraphs[0].text = case["header"]
    if "table" in case:
        rows = case["table"]
        table = doc.add_table(rows=len(rows), cols=len(rows[0]))
        table.style = "Table Grid"
        for row, values in zip(table.rows, rows, strict=True):
            for cell, value in zip(row.cells, values, strict=True):
                cell.text = value
    doc.save(str(target))


def document_contents(path: Path) -> dict:
    doc = Document(str(path))
    return {
        "paragraphs": [paragraph.text for paragraph in doc.paragraphs],
        "tables": [[[cell.text for cell in row.cells] for row in table.rows] for table in doc.tables],
        "headers": [paragraph.text for paragraph in doc.sections[0].header.paragraphs],
    }


def evaluate(corpus_path: Path, output_dir: Path, mode: str = "fixture") -> dict:
    if mode not in {"fixture", "live"}:
        raise ValueError("Mode must be fixture or live")
    corpus = json.loads(corpus_path.read_text(encoding="utf-8"))
    if not isinstance(corpus, list) or not corpus:
        raise ValueError("Corpus must be a nonempty list")
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    original_factory = field_mapper.ChatOpenAI
    totals = {"input_tokens": 0, "output_tokens": 0, "api_calls": 0}
    results = []
    started = time.perf_counter()
    for case in corpus:
        case_dir = output_dir / case["id"]
        case_dir.mkdir(parents=True, exist_ok=True)
        template = case_dir / "before.docx"
        build_template(case, template)
        data = case_dir / "input.json"
        data.write_text(json.dumps(case["data"], ensure_ascii=False, indent=2), encoding="utf-8")
        _, fields, _ = get_all_placeholders(Document(template))
        # Decisions are authored separately from expected document text. Bind only
        # the positional identity here; do not derive expected values from output.
        replies = []
        if fields:
            replies.append(
                json.dumps(
                    [
                        {
                            "id": placeholder_id(fields[item["field"]]),
                            "value": item["value"],
                            "used_json_keys": item["keys"],
                        }
                        for item in case.get("decisions", [])
                    ]
                )
            )
        replies.extend(json.dumps(reply) for reply in case.get("table_replies", []))
        replay = iter(replies)

        class ObservedModel:
            def __init__(self, _replay=replay, **kwargs):
                self.fixture_responses = _replay
                self.model = original_factory(**kwargs) if mode == "live" else None

            def invoke(self, messages):
                if mode == "fixture":
                    return SimpleNamespace(content=next(self.fixture_responses))
                assert self.model is not None
                response = self.model.invoke(messages)
                totals["api_calls"] += 1
                usage = getattr(response, "usage_metadata", None) or {}
                totals["input_tokens"] += usage.get("input_tokens", 0)
                totals["output_tokens"] += usage.get("output_tokens", 0)
                return response

        error = None
        with TemporaryDirectory(prefix="docx-evaluation-") as temporary:
            previous = Path.cwd()
            try:
                os.chdir(temporary)  # Every case gets a cold, independent cache.
                with patch.object(field_mapper, "ChatOpenAI", ObservedModel), contextlib.redirect_stdout(io.StringIO()):
                    run_pipeline(template, data, case_dir / "after.docx")
            except Exception as caught:
                # Never store provider exception messages or credential-bearing URLs.
                error = type(caught).__name__
            finally:
                os.chdir(previous)
        differences: list[str | dict] = []
        if error:
            differences.append(f"Pipeline failed: {error}")
        else:
            actual = document_contents(case_dir / "after.docx")
            for key, expected in case["expected"].items():
                if actual.get(key) != expected:
                    differences.append({"part": key, "expected": expected, "actual": actual.get(key)})
        results.append({"id": case["id"], "passed": not differences, "differences": differences})
    passed = sum(item["passed"] for item in results)
    report = {
        "mode": mode,
        "model": field_mapper._model_name() if mode == "live" else None,
        "cases": len(corpus),
        "passed": passed,
        "document_exact_match_rate": passed / len(corpus),
        "live_model_accuracy": passed / len(corpus) if mode == "live" else None,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
        **totals,
        "estimated_cost_usd": None,
        "results": results,
    }
    (output_dir / "results.json").write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mode", choices=["fixture", "live"], default="fixture")
    parser.add_argument("--allow-api", action="store_true", help="Explicitly authorize billable live model calls")
    parser.add_argument("--corpus", type=Path, default=Path(__file__).with_name("corpus.json"))
    parser.add_argument("--output-dir", type=Path, default=Path("evaluation/output"))
    args = parser.parse_args()
    if args.mode == "live" and not args.allow_api:
        parser.error("Live mode requires --allow-api; synthetic source text will be sent to the configured model")
    load_dotenv()
    result = evaluate(args.corpus, args.output_dir, args.mode)
    print(json.dumps({key: value for key, value in result.items() if key != "results"}, indent=2))
    if result["passed"] != result["cases"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
