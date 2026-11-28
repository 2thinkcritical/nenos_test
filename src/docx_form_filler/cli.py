"""Command-line entry point: ``docx-fill TEMPLATE.docx DATA.json``."""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import NoReturn

from dotenv import load_dotenv
from openai import OpenAIError

from docx_form_filler.pipeline import InputError, run_pipeline
from docx_form_filler.reporting import format_summary


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-fill", description="Fill a DOCX form using JSON data and semantic field mapping."
    )
    parser.add_argument("docx_path", type=Path, help="path to the source DOCX template")
    parser.add_argument("json_path", type=Path, help="path to a JSON object with source data")
    parser.add_argument("--output", type=Path, help="output DOCX path (default: <template>.filled.docx beside it)")
    parser.add_argument("-v", "--verbose", action="store_true", help="show progress details")
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING, format="%(levelname)s %(name)s: %(message)s"
    )
    load_dotenv()
    try:
        result = run_pipeline(args.docx_path, args.json_path, args.output)
    except OpenAIError:
        # Provider messages can include request details; keep them out of the terminal.
        _fail(
            "Model request failed. Check your API key, model access, network, and usage limits. "
            "No completed document was saved."
        )
    except (InputError, OSError) as error:
        _fail(str(error))
    print(format_summary(result.summary, result.report_path))
    print(f"Saved filled document: {result.output_path}")


def _fail(message: str) -> NoReturn:
    print(f"Error: {message}", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()
