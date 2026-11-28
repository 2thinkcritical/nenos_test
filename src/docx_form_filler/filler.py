"""Write mapped values into a DOCX while keeping its formatting.

Text is replaced inside existing runs where possible. Blanks split across runs,
multi-line values and table growth fall back to targeted XML operations.
"""

from __future__ import annotations

import dataclasses
import json
import logging
from collections.abc import Mapping, Sequence
from copy import deepcopy
from dataclasses import dataclass
from typing import Any

from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml.etree import _Element

from docx_form_filler.analyzer import (
    CELL_MARKER_RE,
    HEADER_FOOTER_ATTRIBUTES,
    PlaceholderInfo,
    Story,
    TableInfo,
    paragraph_runs,
    placeholder_id,
)
from docx_form_filler.mapper import FieldMapping
from docx_form_filler.matching import match_columns_by_header, match_table_to_array

logger = logging.getLogger(__name__)

CHECKBOX = "|_|"
CHECKED_BOX = "|X|"


@dataclass
class FillResult:
    filled_count: int = 0
    table_stats: list[dict[str, Any]] = dataclasses.field(default_factory=list)
    table_used_keys: set[str] = dataclasses.field(default_factory=set)
    log: list[dict[str, Any]] = dataclasses.field(default_factory=list)


def array_inputs(json_data: Mapping[str, Any]) -> dict[str, list[Any]]:
    """Top-level values that are arrays, or strings that decode to arrays."""
    arrays: dict[str, list[Any]] = {}
    for key, value in json_data.items():
        if isinstance(value, str):
            try:
                value = json.loads(value)
            except ValueError:
                continue
        if isinstance(value, list):
            arrays[key] = value
    return arrays


def exclude_array_placeholders(
    placeholders: Sequence[PlaceholderInfo],
    data_tables: Sequence[TableInfo],
    json_data: Mapping[str, Any],
    table_key_mapping: Mapping[int, str] | None = None,
) -> list[PlaceholderInfo]:
    """Drop blanks in rows owned by array data, so each cell has a single writer."""
    arrays = array_inputs(json_data)
    owned: set[tuple[int, int]] = set()
    for table in data_tables:
        if _array_key_for(table, arrays, table_key_mapping or {}) is not None:
            owned.update((table.table_index, row) for row in table.data_row_indices)
    return [
        field
        for field in placeholders
        if not (
            field.location_type == "table"
            and field.header_footer_type == "body"
            and (field.table_index, field.row_index) in owned
        )
    ]


def fill_document(
    doc: Document,
    placeholders: Sequence[PlaceholderInfo],
    mapping: FieldMapping,
    json_data: Mapping[str, Any],
    data_tables: Sequence[TableInfo] = (),
    table_key_mapping: Mapping[int, str] | None = None,
    column_mappings: Mapping[int, Mapping[int, str]] | None = None,
) -> FillResult:
    """Fill ``doc`` in place: paragraph blanks, table-cell blanks, then array tables."""
    result = FillResult()
    fields = exclude_array_placeholders(placeholders, data_tables, json_data, table_key_mapping)
    _fill_paragraph_placeholders(doc, fields, mapping, result)
    _fill_table_placeholders(doc, fields, mapping, result)
    _fill_data_tables(doc, data_tables, json_data, table_key_mapping or {}, column_mappings or {}, result)
    logger.info("Document filled: %d placeholders replaced", result.filled_count)
    return result


def _array_key_for(table: TableInfo, arrays: Mapping[str, list[Any]], mapping: Mapping[int, str]) -> str | None:
    """The model's choice when valid, otherwise the offline heuristic."""
    key = mapping.get(table.table_index)
    return key if key in arrays else match_table_to_array(table, arrays)


def _story(doc: Document, section_index: int, header_footer_type: str) -> Story | None:
    """The body, or the header/footer a placeholder was found in."""
    if header_footer_type == "body":
        return doc  # Body paragraph indices are absolute within doc.paragraphs.
    attribute = HEADER_FOOTER_ATTRIBUTES.get(header_footer_type)
    if attribute is None or not 0 <= section_index < len(doc.sections):
        return None
    story: Story = getattr(doc.sections[section_index], attribute)
    return story


def _decision(mapping: FieldMapping, field: PlaceholderInfo) -> tuple[str, list[str]] | None:
    decision = mapping.get(placeholder_id(field))
    value = decision.get("value") if decision else None
    if decision is None or value is None:
        return None
    return str(value), decision.get("used_keys", [])


def _fill_paragraph_placeholders(
    doc: Document, placeholders: Sequence[PlaceholderInfo], mapping: FieldMapping, result: FillResult
) -> None:
    grouped: dict[tuple[int, str, int], list[PlaceholderInfo]] = {}
    for field in placeholders:
        if field.location_type == "paragraph":
            grouped.setdefault((field.section_index, field.header_footer_type, field.paragraph_index), []).append(field)

    # Resolve every paragraph before editing: multi-line values insert paragraphs and shift indices.
    paragraphs: dict[tuple[int, str, int], Paragraph] = {}
    for key in grouped:
        story = _story(doc, key[0], key[1])
        if story is not None and key[2] < len(story.paragraphs):
            paragraphs[key] = story.paragraphs[key[2]]

    text_count = checkbox_count = 0
    for key in sorted(grouped):
        paragraph = paragraphs.get(key)
        if paragraph is None:
            continue
        # Right to left, so earlier character offsets stay valid.
        for field in sorted(grouped[key], key=lambda item: item.char_offset, reverse=True):
            decision = _decision(mapping, field)
            if decision is None:
                continue
            value, used_keys = decision
            if field.is_checkbox:
                _fill_checkbox(paragraph, field, value)
                checkbox_count += 1
            elif _fill_text_blank(paragraph, field, value):
                text_count += 1
            result.log.append(
                {
                    "type": "placeholder",
                    "id": placeholder_id(field),
                    "value": value,
                    "source": used_keys,
                    "context": field.context_before + field.placeholder_pattern + field.context_after,
                }
            )
    result.filled_count += text_count + checkbox_count
    logger.info("Filled %d text blanks and %d checkboxes in paragraphs", text_count, checkbox_count)


def _fill_table_placeholders(
    doc: Document, placeholders: Sequence[PlaceholderInfo], mapping: FieldMapping, result: FillResult
) -> None:
    cells = [field for field in placeholders if field.location_type == "table"]
    # Right to left within each cell paragraph, so earlier offsets stay valid.
    cells.sort(
        key=lambda f: (
            f.section_index,
            f.header_footer_type,
            f.table_index,
            f.row_index,
            f.cell_index,
            f.cell_paragraph_index,
            f.char_offset,
        ),
        reverse=True,
    )
    filled = 0
    for field in cells:
        decision = _decision(mapping, field)
        paragraph = _cell_paragraph(doc, field)
        if decision is None or paragraph is None:
            continue
        value, used_keys = decision
        if field.is_checkbox:
            if value not in {"CHECKED", "UNCHECKED"}:
                continue
            _fill_checkbox(paragraph, field, value)
        elif not _fill_text_blank(paragraph, field, value):
            continue
        filled += 1
        result.log.append(
            {
                "type": "placeholder_table",
                "id": placeholder_id(field),
                "value": value,
                "source": used_keys,
                "table_idx": field.table_index,
                "row": field.row_index,
                "col": field.cell_index,
            }
        )
    result.filled_count += filled
    logger.info("Filled %d table cell placeholders", filled)


def _cell_paragraph(doc: Document, field: PlaceholderInfo) -> Paragraph | None:
    story = _story(doc, field.section_index, field.header_footer_type)
    if story is None or field.table_index >= len(story.tables):
        return None
    table = story.tables[field.table_index]
    if field.row_index >= len(table.rows):
        return None
    row = table.rows[field.row_index]
    if field.cell_index >= len(row.cells):
        return None
    paragraphs = row.cells[field.cell_index].paragraphs
    return paragraphs[field.cell_paragraph_index] if 0 <= field.cell_paragraph_index < len(paragraphs) else None


def _fill_data_tables(
    doc: Document,
    data_tables: Sequence[TableInfo],
    json_data: Mapping[str, Any],
    table_key_mapping: Mapping[int, str],
    column_mappings: Mapping[int, Mapping[int, str]],
    result: FillResult,
) -> None:
    arrays = array_inputs(json_data)
    if not arrays:
        return
    for info in data_tables:
        key = _array_key_for(info, arrays, table_key_mapping)
        if info.table_index >= len(doc.tables) or key is None:
            continue
        table, rows = doc.tables[info.table_index], arrays[key]
        result.table_used_keys.add(key)
        logger.info("Filling data table %d from %r (%d rows)", info.table_index, key, len(rows))
        _clear_marker_cells(table, info.data_row_indices)
        # Snapshot before writing values, so cloned overflow rows start empty.
        template = deepcopy(table.rows[info.data_row_indices[-1]]._tr) if info.data_row_indices else None

        filled_rows = 0
        for position, item in enumerate(rows):
            if position < len(info.data_row_indices):
                row_index = info.data_row_indices[position]
            elif template is not None:
                row_index = info.data_row_indices[-1] + position - len(info.data_row_indices) + 1
                _insert_row(table, template, row_index)
            else:
                break
            if row_index >= len(table.rows):
                continue
            row = table.rows[row_index]
            for column, value, source in _fill_row(row, item, info.headers, column_mappings.get(info.table_index)):
                result.log.append(
                    {
                        "type": "table_cell",
                        "table_idx": info.table_index,
                        "row": row_index,
                        "col": column,
                        "value": value,
                        "source": f"{key}[{position}]{source}",
                    }
                )
            filled_rows += 1
        result.table_stats.append(
            {"table_index": info.table_index, "json_key": key, "rows_filled": filled_rows, "total_rows": len(rows)}
        )


def _fill_row(
    row: _Row, item: object, headers: Sequence[str], columns: Mapping[int, str] | None
) -> list[tuple[int, str, str]]:
    """Write one array item into a row; returns (column, value, source suffix) per written cell."""
    if isinstance(item, str):
        for column, cell in enumerate(row.cells):
            if not cell.text.strip():
                _set_cell_text(cell, item)
                return [(column, item, "")]
        return []
    if not isinstance(item, dict):
        return []
    columns = columns or match_columns_by_header(headers, list(item))
    written = []
    for column, cell in enumerate(row.cells):
        if column in columns:
            key = columns[column]
            raw = item.get(key)
            value = "" if raw is None else str(raw)
            _set_cell_text(cell, value)
            written.append((column, value, f".{key}"))
    return written


def _clear_marker_cells(table: Table, row_indices: Sequence[int]) -> None:
    """Empty cells that contain only blank markers, before rows are filled or cloned."""
    for row_index in row_indices:
        for cell in table.rows[row_index].cells:
            if cell.text.strip() and not CELL_MARKER_RE.sub("", cell.text).strip():
                for paragraph in cell.paragraphs:
                    for run in paragraph_runs(paragraph):
                        for text in run._r.findall(qn("w:t")):
                            text.text = ""


def _insert_row(table: Table, template: _Element, index: int) -> None:
    """Clone ``template`` at ``index``, keeping totals and notes after the data region."""
    table.rows[index - 1]._tr.addnext(deepcopy(template))


def _fill_checkbox(paragraph: Paragraph, field: PlaceholderInfo, value: str) -> None:
    if value.strip().upper() == "CHECKED":
        _replace_across_runs(paragraph, CHECKBOX, CHECKED_BOX, field.char_offset if field.char_offset >= 0 else None)


def _fill_text_blank(paragraph: Paragraph, field: PlaceholderInfo, value: str) -> bool:
    """Replace the blank with ``value``; line breaks become new paragraphs."""
    pattern = field.placeholder_pattern
    if field.char_offset >= 0:
        return _replace_across_runs(paragraph, pattern, value, field.char_offset)
    runs = paragraph.runs
    if 0 <= field.run_index < len(runs) and pattern in runs[field.run_index].text:
        run = runs[field.run_index]
        start = run.text.find(pattern)
        _write_into_run(paragraph, run, run.text[:start], _normalize_newlines(value), run.text[start + len(pattern) :])
        return True
    return _replace_across_runs(paragraph, pattern, value)


def _replace_across_runs(paragraph: Paragraph, pattern: str, value: str, start: int | None = None) -> bool:
    """Replace ``pattern`` (at ``start``, or its first occurrence), even when split across runs."""
    runs = paragraph_runs(paragraph)
    text = "".join(run.text for run in runs)
    if start is None:
        start = text.find(pattern)
    if start < 0 or text[start : start + len(pattern)] != pattern:
        return False
    end = start + len(pattern)

    first = last = -1
    first_offset = last_offset = position = 0
    for index, run in enumerate(runs):
        run_end = position + len(run.text)
        if first == -1 and start < run_end:
            first, first_offset = index, start - position
        if first != -1 and end <= run_end:
            last, last_offset = index, end - position
            break
        position = run_end
    if first == -1 or last == -1:
        return True  # Matched text that no run covers; nothing safe to edit.

    value = _normalize_newlines(value)
    before, after = runs[first].text[:first_offset], runs[last].text[last_offset:]
    if first == last or "\n" in value:
        for run in runs[first + 1 : last + 1]:
            run.text = ""
        _write_into_run(paragraph, runs[first], before, value, after)
    else:
        # Keep the trailing text in its own run so it keeps its formatting.
        runs[first].text = before + value
        for run in runs[first + 1 : last]:
            run.text = ""
        runs[last].text = after
    return True


def _write_into_run(paragraph: Paragraph, run: Run, before: str, value: str, after: str) -> None:
    if "\n" in value:
        _split_paragraph_at_newlines(paragraph, run, value, before, after)
    else:
        run.text = before + value + after


def _normalize_newlines(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n")


def _split_paragraph_at_newlines(
    paragraph: Paragraph, run: Run, value: str, text_before: str = "", text_after: str = ""
) -> None:
    """Write a multi-line value as consecutive paragraphs with the original formatting.

    The first line stays in ``run``; later lines become new paragraphs cloned from
    the paragraph and run properties. Runs after ``run`` move to the last new
    paragraph, so text following the blank keeps its position.
    """
    lines = value.split("\n")
    run.text = text_before + lines[0]

    all_runs = paragraph.runs
    index = next((i for i, candidate in enumerate(all_runs) if candidate._r is run._r), -1)
    trailing_runs = all_runs[index + 1 :] if index != -1 else []
    paragraph_properties = paragraph._p.find(qn("w:pPr"))
    run_properties = run._r.find(qn("w:rPr"))

    previous = paragraph._p
    for number, line in enumerate(lines[1:], start=1):
        is_last = number == len(lines) - 1
        new_paragraph = OxmlElement("w:p")
        if paragraph_properties is not None:
            properties = deepcopy(paragraph_properties)
            paragraph_id = properties.find(qn("w:pId"))
            if paragraph_id is not None:
                properties.remove(paragraph_id)  # IDs must stay unique.
            new_paragraph.append(properties)
        new_run = OxmlElement("w:r")
        if run_properties is not None:
            new_run.append(deepcopy(run_properties))
        text = OxmlElement("w:t")
        text.text = line + text_after if is_last else line
        if text.text and (text.text[0] == " " or text.text[-1] == " "):
            text.set(qn("xml:space"), "preserve")
        new_run.append(text)
        new_paragraph.append(new_run)
        if is_last:
            for trailing in trailing_runs:
                new_paragraph.append(trailing._r)  # lxml moves the element.
        previous.addnext(new_paragraph)
        previous = new_paragraph


def _set_cell_text(cell: _Cell, text: str) -> None:
    """Replace a cell's text, keeping the first run's formatting."""
    text = _normalize_newlines(text)
    if not cell.paragraphs:
        cell.text = text
        return
    paragraph = cell.paragraphs[0]
    runs = paragraph.runs
    for extra in runs[1:]:
        extra.text = ""
    run = runs[0] if runs else paragraph.add_run("")
    if "\n" in text:
        _split_paragraph_at_newlines(paragraph, run, text)
    else:
        run.text = text
