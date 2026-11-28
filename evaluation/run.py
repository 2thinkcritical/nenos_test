"""Evaluate saved DOCX contents against independent expected outputs.

Fixture mode checks application behavior with authored model responses. It does
not measure model accuracy. Live mode is opt-in and makes billable API requests.
"""

import argparse
import contextlib
import io
import json
import os
import re
import time
from collections.abc import Iterator
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from typing import Any
from unittest.mock import patch

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import RGBColor
from dotenv import load_dotenv
from langchain_core.messages import BaseMessage

from docx_form_filler import llm
from docx_form_filler.analyzer import analyze_document, placeholder_id
from docx_form_filler.pipeline import run_pipeline

CASE_ID_RE = re.compile(r"[a-z0-9-]{1,80}")
EXPECTED_PARTS = {"paragraphs", "tables", "headers"}


def build_template(case: dict[str, Any], target: Path) -> None:
    doc = Document()
    doc.core_properties.author = "Synthetic evaluation fixture"
    doc.core_properties.last_modified_by = "Synthetic evaluation fixture"
    title_style = doc.styles["Title"]
    title_style.font.color.rgb = RGBColor(0, 0, 0)
    for border in title_style.element.xpath("./w:pPr/w:pBdr"):
        border.getparent().remove(border)
    for index, text in enumerate(case["paragraphs"]):
        paragraph = doc.add_paragraph()
        if index == 0:
            paragraph.style = title_style
            for border in paragraph._p.xpath("./w:pPr/w:pBdr"):
                border.getparent().remove(border)
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


def document_contents(path: Path) -> dict[str, Any]:
    doc = Document(str(path))
    return {
        "paragraphs": [paragraph.text for paragraph in doc.paragraphs],
        "tables": [[[cell.text for cell in row.cells] for row in table.rows] for table in doc.tables],
        "headers": [paragraph.text for paragraph in doc.sections[0].header.paragraphs],
    }


def _is_text_list(value: object) -> bool:
    return isinstance(value, list) and all(isinstance(text, str) for text in value)


def _is_table(value: object) -> bool:
    return (
        isinstance(value, list)
        and bool(value)
        and isinstance(value[0], list)
        and all(_is_text_list(row) and len(row) == len(value[0]) and bool(row) for row in value)
    )


def _validate_output_path(path: Path, root: Path, corpus_path: Path, *, directory: bool = False) -> None:
    if path.is_symlink() or not path.resolve().is_relative_to(root):
        raise ValueError(f"Unsafe evaluation output path: {path}")
    if path.exists():
        if directory and path.is_dir():
            return
        if not directory and path.is_file() and path.stat().st_nlink == 1 and not path.samefile(corpus_path):
            return
        raise ValueError(f"Conflicting evaluation output path: {path}")


def _validate_corpus(corpus: object, corpus_path: Path, output_dir: Path) -> None:
    if not isinstance(corpus, list) or not corpus:
        raise ValueError("Corpus must be a nonempty list")
    _validate_output_path(output_dir, output_dir, corpus_path, directory=True)
    _validate_output_path(output_dir / "results.json", output_dir, corpus_path)
    ids: set[str] = set()
    for index, case in enumerate(corpus, start=1):
        if not isinstance(case, dict):
            raise ValueError(f"Corpus case {index} must be an object")
        case_id = case.get("id")
        if not isinstance(case_id, str) or not CASE_ID_RE.fullmatch(case_id):
            raise ValueError(f"Corpus case {index} ID must use 1-80 lowercase letters, digits, or hyphens")
        if case_id in ids:
            raise ValueError(f"Duplicate corpus case ID: {case_id}")
        ids.add(case_id)

        expected = case.get("expected")
        if not isinstance(expected, dict) or not expected or expected.keys() - EXPECTED_PARTS:
            raise ValueError(f"Corpus case {case_id} must expect paragraphs, tables, or headers")
        for part, value in expected.items():
            valid = (
                isinstance(value, list) and all(_is_table(table) for table in value)
                if part == "tables"
                else _is_text_list(value)
            )
            if not valid:
                raise ValueError(f"Corpus case {case_id} has invalid expected {part}")
        if not any(expected.values()):
            raise ValueError(f"Corpus case {case_id} must include at least one nonempty expected part")

        if not _is_text_list(case.get("paragraphs")) or not isinstance(case.get("data"), dict):
            raise ValueError(f"Corpus case {case_id} must provide a paragraphs list and data object")
        for part in ("header", "hyperlink"):
            if part in case and not isinstance(case[part], str):
                raise ValueError(f"Corpus case {case_id} has invalid {part} text")
        if "table" in case and not _is_table(case["table"]):
            raise ValueError(f"Corpus case {case_id} must provide a nonempty rectangular table of strings")

        case_dir = output_dir / case_id
        _validate_output_path(case_dir, output_dir, corpus_path, directory=True)
        for name in ("before.docx", "input.json", "after.docx"):
            _validate_output_path(case_dir / name, output_dir, corpus_path)


def evaluate(corpus_path: Path, output_dir: Path, mode: str = "fixture") -> dict[str, Any]:
    if mode not in {"fixture", "live"}:
        raise ValueError("Mode must be fixture or live")
    corpus = json.loads(corpus_path.read_text(encoding="utf-8"))
    if output_dir.is_symlink():
        raise ValueError(f"Unsafe evaluation output path: {output_dir}")
    output_dir = output_dir.resolve()
    _validate_corpus(corpus, corpus_path, output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    original_factory = llm.ChatOpenAI  # captured before the per-case patch below
    api_calls = successful_responses = responses_with_usage = 0
    input_tokens = output_tokens = 0
    results = []
    started = time.perf_counter()
    for case in corpus:
        case_dir = output_dir / case["id"]
        case_dir.mkdir(parents=True, exist_ok=True)
        after = case_dir / "after.docx"
        after.unlink(missing_ok=True)
        template = case_dir / "before.docx"
        build_template(case, template)
        data = case_dir / "input.json"
        data.write_text(json.dumps(case["data"], ensure_ascii=False, indent=2), encoding="utf-8")
        # Decisions are authored separately from expected document text. Bind only
        # the positional identity here; do not derive expected values from output.
        replies = []
        if mode == "fixture":
            fields, _ = analyze_document(Document(template))
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
        responses_before_case = successful_responses
        failure_stage = None

        class ObservedModel:
            def __init__(self, _replay: Iterator[str] = replay, **kwargs: Any) -> None:
                nonlocal failure_stage
                self.fixture_responses = _replay
                try:
                    self.model = original_factory(**kwargs) if mode == "live" else None
                except Exception:
                    failure_stage = "model_initialization"
                    raise

            def invoke(self, messages: list[BaseMessage]) -> Any:
                nonlocal \
                    api_calls, \
                    successful_responses, \
                    responses_with_usage, \
                    input_tokens, \
                    output_tokens, \
                    failure_stage
                if mode == "fixture":
                    return SimpleNamespace(content=next(self.fixture_responses))
                assert self.model is not None
                api_calls += 1
                try:
                    response = self.model.invoke(messages)
                except Exception:
                    failure_stage = "model_request"
                    raise
                successful_responses += 1
                usage = getattr(response, "usage_metadata", None) or {}
                request_input = usage.get("input_tokens")
                request_output = usage.get("output_tokens")
                if (
                    type(request_input) is int
                    and type(request_output) is int
                    and min(request_input, request_output) >= 0
                ):
                    input_tokens += request_input
                    output_tokens += request_output
                    responses_with_usage += 1
                return response

        differences: list[str | dict[str, Any]] = []
        with TemporaryDirectory(prefix="docx-evaluation-") as temporary:
            previous = Path.cwd()
            try:
                os.chdir(temporary)  # Every case gets a cold, independent cache.
                with patch.object(llm, "ChatOpenAI", ObservedModel), contextlib.redirect_stdout(io.StringIO()):
                    run_pipeline(template, data, after)
                actual = document_contents(after)
            except Exception as caught:
                # Never store provider exception messages or credential-bearing URLs.
                failure_stage = failure_stage or "pipeline"
                differences.append(f"Pipeline failed: {type(caught).__name__}")
            else:
                if mode == "live" and successful_responses == responses_before_case:
                    failure_stage = "no_model_response"
                for key, expected in case["expected"].items():
                    if actual[key] != expected:
                        differences.append({"part": key, "expected": expected, "actual": actual[key]})
            finally:
                os.chdir(previous)
        scored = mode == "fixture" or failure_stage is None
        results.append(
            {
                "id": case["id"],
                "scored": scored,
                "passed": scored and not differences,
                "failure_stage": failure_stage,
                "differences": differences,
            }
        )
    passed = sum(item["passed"] for item in results)
    scored_cases = sum(item["scored"] for item in results)
    exact_match_rate = passed / scored_cases if scored_cases else None
    usage_complete = mode == "fixture" or (api_calls > 0 and responses_with_usage == api_calls)
    report = {
        "mode": mode,
        "model": llm.model_name() if mode == "live" else None,
        "cases": len(corpus),
        "passed": passed,
        "scored_cases": scored_cases,
        "unscored_cases": len(corpus) - scored_cases,
        "document_exact_match_rate": exact_match_rate,
        "live_model_accuracy": exact_match_rate if mode == "live" else None,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
        "api_calls": api_calls,
        "successful_model_responses": successful_responses,
        "input_tokens": input_tokens if usage_complete else None,
        "output_tokens": output_tokens if usage_complete else None,
        "usage_complete": usage_complete,
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
