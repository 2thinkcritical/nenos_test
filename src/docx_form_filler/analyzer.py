"""Locate fillable blanks and data tables in a DOCX document.

Recognised visual markers:
  - underscores: ``___`` (3+)
  - dots: ``....`` (4+)
  - Unicode ellipses: ``…`` (1+, optionally mixed with a dot)
  - text checkboxes: ``|_|``

Each blank is described by a :class:`PlaceholderInfo` carrying enough surrounding
text for the semantic mapper and enough coordinates for the filler.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Protocol

from docx.document import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run

logger = logging.getLogger(__name__)

BLANK_RE = re.compile(r"_{3,}|\.{4,}|…+\.?…*")
CHECKBOX_RE = re.compile(r"\|_\|")
CELL_MARKER_RE = re.compile(rf"{BLANK_RE.pattern}|{CHECKBOX_RE.pattern}")

LocationType = Literal["paragraph", "table"]


class Story(Protocol):
    """Anything with paragraphs and tables: the body, a header/footer, or a table cell."""

    @property
    def paragraphs(self) -> list[Paragraph]: ...

    @property
    def tables(self) -> list[Table]: ...


#: Header/footer story name -> attribute on ``docx.section.Section``.
HEADER_FOOTER_ATTRIBUTES = {
    "header_primary": "header",
    "header_first": "first_page_header",
    "header_even": "even_page_header",
    "footer_primary": "footer",
    "footer_first": "first_page_footer",
    "footer_even": "even_page_footer",
}


@dataclass
class PlaceholderInfo:
    """A single fillable blank or checkbox found in the document."""

    location_type: LocationType
    paragraph_index: int  # index in the container's paragraphs; -1 for table cells
    run_index: int  # run containing the start of the blank
    placeholder_pattern: str  # the literal marker text, e.g. "______"
    context_before: str
    context_after: str
    full_paragraph_text: str
    char_offset: int = -1  # offset in the paragraph; distinguishes repeated blanks
    surrounding_context: str = ""  # nearby paragraphs, for broader semantic context
    table_index: int = -1
    row_index: int = -1
    cell_index: int = -1
    cell_paragraph_index: int = -1
    column_header: str = ""
    is_checkbox: bool = False
    checkbox_label: str = ""
    section_index: int = -1
    header_footer_type: str = "body"  # "body" or a key of HEADER_FOOTER_ATTRIBUTES


@dataclass
class TableInfo:
    """A table whose empty rows should receive array data."""

    table_index: int
    headers: list[str]
    data_row_indices: list[int]  # first contiguous range of fillable rows after the header
    context: str
    section_index: int = -1
    header_footer_type: str = "body"


@dataclass(frozen=True)
class _Location:
    """Where a scanned paragraph lives; shared by every blank found in it."""

    section_index: int
    header_footer_type: str
    location_type: LocationType = "paragraph"
    paragraph_index: int = -1
    table_index: int = -1
    row_index: int = -1
    cell_index: int = -1
    cell_paragraph_index: int = -1
    column_header: str = ""


def placeholder_id(p: PlaceholderInfo) -> str:
    """Stable identity shared by the mapper and the filler."""
    prefix = f"s{p.section_index}_{p.header_footer_type}" if p.header_footer_type != "body" else "body"
    if p.location_type == "table":
        base = f"{prefix}_table_{p.table_index}_r{p.row_index}_c{p.cell_index}_p{p.cell_paragraph_index}"
    else:
        base = f"{prefix}_para_{p.paragraph_index}_run{p.run_index}_{p.placeholder_pattern[:10]}"
    return f"{base}_at{p.char_offset}" if p.char_offset >= 0 else base


def analyze_document(doc: Document) -> tuple[list[PlaceholderInfo], list[TableInfo]]:
    """Return every placeholder (body, tables, headers, footers) and the data tables."""
    placeholders: list[PlaceholderInfo] = []
    _scan_container(doc, placeholders, section_index=-1, header_footer_type="body")
    for section_index, section in enumerate(doc.sections):
        for story, attribute in HEADER_FOOTER_ATTRIBUTES.items():
            container = getattr(section, attribute)
            if container and not container.is_linked_to_previous:
                _scan_container(container, placeholders, section_index=section_index, header_footer_type=story)
    return placeholders, extract_data_tables(doc)


def paragraph_runs(paragraph: Paragraph) -> list[Run]:
    """Text runs in display order, including runs inside hyperlinks."""
    runs: list[Run] = []
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            runs.extend(item.runs)
        else:
            runs.append(item)
    return runs


def extract_data_tables(doc: Document) -> list[TableInfo]:
    """Identify body tables with empty rows to be filled from arrays.

    With multi-row headers, the row with the most distinct labels wins, which
    selects the per-column header rather than a merged title row.
    """
    tables = []
    for table_index, table in enumerate(doc.tables):
        best_headers: list[str] = []
        best_score = 0
        data_rows: list[int] = []
        for row_index, row in enumerate(table.rows):
            cells = row.cells
            texts = [cell.text.strip() for cell in cells]
            labels = {text for text in texts if CELL_MARKER_RE.sub("", text).strip()}
            if not labels and not any(cell.tables for cell in cells):
                if best_headers:
                    data_rows.append(row_index)
                continue
            if data_rows:
                break  # A total or note ends the fillable region.
            if len(cells) > 1 and all(cell._tc is cells[0]._tc for cell in cells):
                continue  # A title spanning the entire table is not a header.
            if len(labels) > best_score:
                best_score = len(labels)
                best_headers = texts
        if data_rows:
            tables.append(
                TableInfo(
                    table_index=table_index,
                    headers=best_headers,
                    data_row_indices=data_rows,
                    context=_table_context(doc, table_index),
                )
            )
    return tables


def write_analysis_report(placeholders: Sequence[PlaceholderInfo], tables: Sequence[TableInfo], path: Path) -> None:
    """Write every detected placeholder and data table to a plain-text report."""
    lines = ["DOCX ANALYSIS REPORT", "====================", "", f"Total placeholders found: {len(placeholders)}"]
    lines.append("-" * 40)
    for index, p in enumerate(placeholders):
        lines += [
            f"PLACEHOLDER [{index}]",
            "=" * 80,
            f"Location type:      {p.location_type}",
            f"Section index:      {p.section_index}",
            f"Header/footer type: {p.header_footer_type}",
            f"Pattern found:      {p.placeholder_pattern!r}",
            f"Is checkbox:        {p.is_checkbox}",
        ]
        if p.is_checkbox:
            lines.append(f"Checkbox label:     {p.checkbox_label!r}")
        if p.location_type == "paragraph":
            lines += [
                "",
                "--- Paragraph ---",
                f"Paragraph index:    {p.paragraph_index}",
                f"Run index:          {p.run_index}",
                f"Full text:          {p.full_paragraph_text!r}",
                f"Context before:     {p.context_before!r}",
                f"Context after:      {p.context_after!r}",
                f"Surrounding:        {p.surrounding_context!r}",
            ]
        else:
            lines += [
                "",
                "--- Table cell ---",
                f"Table index:        {p.table_index}",
                f"Row index:          {p.row_index}",
                f"Cell index:         {p.cell_index}",
                f"Cell paragraph:     {p.cell_paragraph_index}",
                f"Column header:      {p.column_header!r}",
                f"Full cell text:     {p.full_paragraph_text!r}",
            ]
        lines.append("")
    lines += ["", "=" * 80, f"DATA TABLES ({len(tables)})", "=" * 80, ""]
    for table in tables:
        lines += [
            f"Table {table.table_index}:",
            f"Headers: {table.headers}",
            f"Data row indices: {table.data_row_indices}",
            f"Context: {table.context}",
            "-" * 40,
        ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    logger.info("Analysis report saved to %s", path.resolve())


def _scan_container(
    container: Story,
    placeholders: list[PlaceholderInfo],
    section_index: int,
    header_footer_type: str,
) -> None:
    """Scan a story (body, header, footer or cell) and its tables, recursing into nested tables."""
    paragraphs = container.paragraphs
    for paragraph_index, paragraph in enumerate(paragraphs):
        if paragraph.text.strip():
            location = _Location(section_index, header_footer_type, paragraph_index=paragraph_index)
            placeholders.extend(_scan_paragraph(paragraph, paragraphs, paragraph_index, location))

    for table_index, table in enumerate(container.tables):
        headers = next(
            (
                texts
                for texts in ([cell.text.strip() for cell in row.cells] for row in table.rows)
                if any(text and not BLANK_RE.fullmatch(text) for text in texts)
            ),
            [],
        )
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                for cell_paragraph_index, paragraph in enumerate(cell.paragraphs):
                    location = _Location(
                        section_index,
                        header_footer_type,
                        location_type="table",
                        table_index=table_index,
                        row_index=row_index,
                        cell_index=cell_index,
                        cell_paragraph_index=cell_paragraph_index,
                        column_header=headers[cell_index] if cell_index < len(headers) else "",
                    )
                    placeholders.extend(_scan_paragraph(paragraph, cell.paragraphs, cell_paragraph_index, location))
                if cell.tables:
                    _scan_container(cell, placeholders, section_index, header_footer_type)


def _scan_paragraph(
    paragraph: Paragraph, siblings: Sequence[Paragraph], index: int, location: _Location
) -> Iterator[PlaceholderInfo]:
    """Yield checkboxes first, then text blanks, in the order they appear."""
    text = paragraph.text
    surrounding = _surrounding_context(siblings, index, text)
    for is_checkbox, pattern in ((True, CHECKBOX_RE), (False, BLANK_RE)):
        for match in pattern.finditer(text):
            after = text[match.end() :].strip()
            yield PlaceholderInfo(
                location_type=location.location_type,
                paragraph_index=location.paragraph_index,
                run_index=_run_at(paragraph, match.start()),
                placeholder_pattern=match.group(),
                char_offset=match.start(),
                context_before=text[: match.start()].strip(),
                context_after=after,
                full_paragraph_text=text,
                surrounding_context=surrounding,
                table_index=location.table_index,
                row_index=location.row_index,
                cell_index=location.cell_index,
                cell_paragraph_index=location.cell_paragraph_index,
                column_header=location.column_header,
                is_checkbox=is_checkbox,
                checkbox_label=after.split("\n")[0].strip() if is_checkbox else "",
                section_index=location.section_index,
                header_footer_type=location.header_footer_type,
            )


def _surrounding_context(paragraphs: Sequence[Paragraph], index: int, text: str) -> str:
    """Text of up to two paragraphs on either side, for broader context."""
    parts = []
    for offset in (-2, -1, 1, 2):
        neighbor = index + offset
        if 0 <= neighbor < len(paragraphs):
            neighbor_text = paragraphs[neighbor].text.strip()
            if neighbor_text and neighbor_text != text.strip():
                parts.append(neighbor_text[:150])
    return " | ".join(parts)


def _run_at(paragraph: Paragraph, char_pos: int) -> int:
    """Index of the run containing ``char_pos`` (the last run if past the end)."""
    runs = paragraph_runs(paragraph)
    position = 0
    for run_index, run in enumerate(runs):
        position += len(run.text)
        if position > char_pos:
            return run_index
    return max(0, len(runs) - 1)


def _table_context(doc: Document, table_index: int) -> str:
    """Up to three non-empty body paragraphs preceding the given table."""
    tables_seen = 0
    context: list[str] = []
    for element in doc.element.body:
        if element.tag == qn("w:tbl"):
            if tables_seen == table_index:
                break
            tables_seen += 1
        elif element.tag == qn("w:p"):
            text = "".join(node.text or "" for node in element.iter(qn("w:t"))).strip()
            if text:
                context = [*context, text][-3:]
    return " | ".join(context)
