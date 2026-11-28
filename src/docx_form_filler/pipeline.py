"""The analyze -> map -> fill -> save pipeline."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any
from zipfile import BadZipFile

import docx
from docx.document import Document
from docx.opc.exceptions import PackageNotFoundError
from lxml.etree import XMLSyntaxError

from docx_form_filler.analyzer import TableInfo, analyze_document, write_analysis_report
from docx_form_filler.filler import array_inputs, exclude_array_placeholders, fill_document
from docx_form_filler.mapper import map_fields, map_table_columns, map_tables_to_arrays, object_keys
from docx_form_filler.reporting import FillSummary, append_report, summarize

logger = logging.getLogger(__name__)

REPORT_PATH = Path("docx_full_log.txt")


class InputError(ValueError):
    """An input or output path the user must correct."""


@dataclass
class PipelineResult:
    output_path: Path
    report_path: Path
    summary: FillSummary


def default_output_path(template: Path) -> Path:
    return template.with_name(f"{template.stem}.filled.docx")


def run_pipeline(
    docx_path: Path, json_path: Path, output_path: Path | None = None, report_path: Path = REPORT_PATH
) -> PipelineResult:
    """Fill a copy of ``docx_path``; the template itself is never modified."""
    output_path = output_path or default_output_path(docx_path)
    _check_paths(docx_path, json_path, output_path)
    doc = _load_docx(docx_path)
    json_data = _load_json(json_path)

    placeholders, data_tables = analyze_document(doc)
    logger.info("Found %d placeholders and %d data tables", len(placeholders), len(data_tables))
    write_analysis_report(placeholders, data_tables, report_path)

    mapping = map_fields(placeholders, json_data)
    table_keys, column_mappings = _map_data_tables(data_tables, json_data)
    placeholders = exclude_array_placeholders(placeholders, data_tables, json_data, table_keys)
    fill = fill_document(doc, placeholders, mapping, json_data, data_tables, table_keys, column_mappings)

    summary = summarize(placeholders, data_tables, mapping, list(json_data), fill)
    append_report(summary, report_path)
    _save_atomically(doc, output_path)
    logger.info("Saved filled document to %s", output_path)
    return PipelineResult(output_path, report_path, summary)


def _check_paths(docx_path: Path, json_path: Path, output_path: Path) -> None:
    if output_path.resolve() == docx_path.resolve() or (
        output_path.exists() and docx_path.exists() and output_path.samefile(docx_path)
    ):
        raise InputError("Output must not overwrite the input DOCX template")
    if output_path.suffix.lower() != ".docx":
        raise InputError("Output path must end with .docx")
    if not docx_path.is_file():
        raise InputError(f"DOCX file not found: {docx_path}")
    if not json_path.is_file():
        raise InputError(f"JSON file not found: {json_path}")


def _load_docx(path: Path) -> Document:
    try:
        return docx.Document(str(path))
    except (PackageNotFoundError, BadZipFile, KeyError, XMLSyntaxError, ValueError) as error:
        raise InputError(f"Cannot read DOCX: {path}. Provide a valid Word .docx file.") from error


def _load_json(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise InputError(f"Invalid JSON input: {error}") from error
    if not isinstance(data, dict):
        raise InputError("JSON root must be an object")
    return data


def _map_data_tables(
    data_tables: list[TableInfo], json_data: dict[str, Any]
) -> tuple[dict[int, str], dict[int, dict[int, str]]]:
    """Ask the model which array fills each table, then which key fills each column."""
    arrays = array_inputs(json_data)
    if not data_tables or not arrays:
        return {}, {}
    table_keys = map_tables_to_arrays(data_tables, arrays)
    column_mappings = {}
    for table in data_tables:
        rows = arrays.get(table_keys.get(table.table_index, ""))
        if rows and isinstance(rows[0], dict):
            column_mappings[table.table_index] = map_table_columns(table.table_index, table.headers, object_keys(rows))
    return table_keys, column_mappings


def _save_atomically(doc: Document, output_path: Path) -> None:
    """Write via a temporary sibling so a failed save never leaves a partial file."""
    target = output_path.resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=target.parent, prefix=".docx-fill-", suffix=".tmp", delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        doc.save(str(temporary_path))
        temporary_path.replace(target)
    finally:
        temporary_path.unlink(missing_ok=True)
