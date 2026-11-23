"""
DOCX Analyzer — Extracts all placeholder positions from a DOCX form.

Scans paragraphs and tables for visual blanks:
  - Underscores: _____ (3+)
  - Dots: ...... (3+)
  - Ellipsis: ………… (1+ unicode ellipsis)
  - Checkboxes: |_|

Returns structured PlaceholderInfo objects with context for LLM mapping.
"""

import os
import re
from dataclasses import dataclass

from docx.document import Document
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# Patterns that indicate a fillable blank
UNDERSCORE_RE = re.compile(r"_{3,}")
DOT_RE = re.compile(r"\.{4,}")
ELLIPSIS_RE = re.compile(r"…+\.?…*")  # unicode ellipsis chars, possibly mixed with dots
CHECKBOX_RE = re.compile(r"\|_\|")

# Combined pattern for any blank
ANY_BLANK_RE = re.compile(r"_{3,}|\.{4,}|…+\.?…*")


@dataclass
class PlaceholderInfo:
    """Represents a single fillable placeholder found in the DOCX."""

    location_type: str  # "paragraph" or "table"
    paragraph_index: int  # index in doc.paragraphs (or -1 for table cells)
    run_index: int  # index of the run containing the blank (-1 if full paragraph)
    placeholder_pattern: str  # the actual blank text (e.g. "___________")
    context_before: str  # text before the blank in the same paragraph
    context_after: str  # text after the blank in the same paragraph
    full_paragraph_text: str  # complete paragraph text for broader context
    char_offset: int = -1  # original paragraph offset; distinguishes repeated blanks
    surrounding_context: str = ""  # text from nearby paragraphs (before + after)
    # Table-specific
    table_index: int = -1
    row_index: int = -1
    cell_index: int = -1
    cell_paragraph_index: int = -1
    column_header: str = ""
    # Checkbox-specific
    is_checkbox: bool = False
    checkbox_label: str = ""
    # Header/Footer/Section awareness
    section_index: int = -1
    header_footer_type: str = "body"  # "body", "header_primary", "header_first", "header_even", "footer_primary", "footer_first", "footer_even"


@dataclass
class TableInfo:
    """Represents a data table that needs to be filled with array data."""

    table_index: int
    headers: list[str]
    data_row_indices: list[int]  # row indices that are empty (data rows)
    context: str  # surrounding text for LLM context
    section_index: int = -1
    header_footer_type: str = "body"


def placeholder_id(p: PlaceholderInfo) -> str:
    """Identity shared by the mapper and filler, including repeated blank positions."""
    prefix = f"s{p.section_index}_{p.header_footer_type}" if p.header_footer_type != "body" else "body"
    if p.location_type == "table":
        base = f"{prefix}_table_{p.table_index}_r{p.row_index}_c{p.cell_index}_p{p.cell_paragraph_index}"
    else:
        base = f"{prefix}_para_{p.paragraph_index}_run{p.run_index}_{p.placeholder_pattern[:10]}"
    return f"{base}_at{p.char_offset}" if p.char_offset >= 0 else base


def _get_surrounding_context(all_paragraphs, p_idx: int, text: str) -> str:
    """Get text from 2 paragraphs before and 2 after for broader context."""
    parts = []
    for offset in [-2, -1, 1, 2]:
        neighbor_idx = p_idx + offset
        if 0 <= neighbor_idx < len(all_paragraphs):
            neighbor_text = all_paragraphs[neighbor_idx].text.strip()
            if neighbor_text and neighbor_text != text.strip():
                parts.append(neighbor_text[:150])
    return " | ".join(parts)


def extract_placeholders(doc: Document) -> list[PlaceholderInfo]:
    """Legacy extraction - redirected to recursive."""
    # This function now needs to return only placeholders, so it will call the new combined function
    # and extract only the placeholders.
    _, placeholders, _ = extract_placeholders_recursive(doc)
    return placeholders


def extract_placeholders_recursive(doc: Document) -> tuple[Document, list[PlaceholderInfo], list[TableInfo]]:
    """
    Recursively extract all placeholders from the document, including:
    - Body paragraphs
    - Tables (and nested tables)
    - Headers
    - Footers

    Also extracts data table information.
    """
    placeholders: list[PlaceholderInfo] = []

    # 1. Body
    # The main document body is effectively "section -1" or generic body
    # But usually it's better to treat it as "body" and maybe associate with section 0?
    # doc.paragraphs covers the main story.
    _scan_container(doc, placeholders, section_index=-1, header_footer_type="body")

    # 2. Key sections (Headers/Footers)
    for s_idx, section in enumerate(doc.sections):
        targets = [
            (section.header, "header_primary"),
            (section.first_page_header, "header_first"),
            (section.even_page_header, "header_even"),
            (section.footer, "footer_primary"),
            (section.first_page_footer, "footer_first"),
            (section.even_page_footer, "footer_even"),
        ]

        for container, hf_type in targets:
            if container and not container.is_linked_to_previous:
                _scan_container(container, placeholders, section_index=s_idx, header_footer_type=hf_type)

    data_tables = extract_data_tables(doc)  # Extract data tables after scanning all placeholders
    return doc, placeholders, data_tables


def save_analysis_report(
    placeholders: list["PlaceholderInfo"], table_infos: list["TableInfo"], filename="docx_structure_debug.txt"
):
    """
    Saves a detailed report of all found placeholders and tables to a text file.
    """
    with open(filename, "w", encoding="utf-8") as f:
        f.write("DOCX ANALYSIS REPORT\n")
        f.write("====================\n\n")

        f.write(f"Total Placeholders Found: {len(placeholders)}\n")
        f.write("-" * 40 + "\n")

        for i, p in enumerate(placeholders):
            f.write(f"PLACEHOLDER [{i}]\n")
            f.write("=" * 80 + "\n")
            f.write(f"Location Type:      {p.location_type}\n")
            f.write(f"Section Index:      {p.section_index}\n")
            f.write(f"Header/Footer Type: {p.header_footer_type}\n")
            f.write(f"Pattern Found:      {repr(p.placeholder_pattern)}\n")
            f.write(f"Is Checkbox:        {p.is_checkbox}\n")
            if p.is_checkbox:
                f.write(f"Checkbox Label:     {repr(p.checkbox_label)}\n")

            if p.location_type == "paragraph":
                f.write("\n--- Paragraph Info ---\n")
                f.write(f"Paragraph Index:    {p.paragraph_index}\n")
                f.write(f"Run Index:          {p.run_index}\n")
                f.write(f"Full Para Text:     {repr(p.full_paragraph_text)}\n")
                f.write(f"Context Before:     {repr(p.context_before)}\n")
                f.write(f"Context After:      {repr(p.context_after)}\n")
                f.write(f"Surrounding Ctx:    {repr(p.surrounding_context)}\n")

            elif p.location_type == "table":
                f.write("\n--- Table Info ---\n")
                f.write(f"Table Index:        {p.table_index}\n")
                f.write(f"Row Index:          {p.row_index}\n")
                f.write(f"Cell Index:         {p.cell_index}\n")
                f.write(f"Cell Para Index:    {p.cell_paragraph_index}\n")
                f.write(f"Column Header:      {repr(p.column_header)}\n")
                f.write(f"Full Cell Text:     {repr(p.full_paragraph_text)}\n")

            f.write("\n")

        f.write("\n" + "=" * 80 + "\n")
        f.write(f"DATA TABLES ({len(table_infos)})\n")
        f.write("=" * 80 + "\n\n")

        for t in table_infos:
            f.write(f"Table {t.table_index} (Index {t.table_index}):\n")
            f.write(f"Headers: {t.headers}\n")
            f.write(f"Data row indices: {t.data_row_indices}\n")
            f.write(f"Context: {t.context}\n")
            f.write("-" * 40 + "\n")

    print(f"Full analysis report saved to: {os.path.abspath(filename)}")


def _scan_container(container, placeholders: list[PlaceholderInfo], section_index: int, header_footer_type: str):
    """Recursively scan a container (Document, Header, Footer, Cell) for paragraphs and tables."""
    # Scan Paragraphs
    for p_idx, para in enumerate(container.paragraphs):
        text = para.text
        if not text.strip():
            continue

        # Checkboxes
        for match in CHECKBOX_RE.finditer(text):
            run_idx = _find_run_for_position(para, match.start())
            after = text[match.end() :].strip()
            label = after.split("\n")[0].strip() if after else ""

            placeholders.append(
                PlaceholderInfo(
                    location_type="paragraph",  # Keeping generic type, using fields for specificity
                    paragraph_index=p_idx,
                    run_index=run_idx,
                    placeholder_pattern=match.group(),
                    char_offset=match.start(),
                    context_before=text[: match.start()].strip(),
                    context_after=text[match.end() :].strip(),
                    full_paragraph_text=text,
                    is_checkbox=True,
                    checkbox_label=label,
                    surrounding_context=_get_surrounding_context(container.paragraphs, p_idx, text),
                    section_index=section_index,
                    header_footer_type=header_footer_type,
                )
            )

        # Text Blanks
        for match in ANY_BLANK_RE.finditer(text):
            run_idx = _find_run_for_position(para, match.start())

            placeholders.append(
                PlaceholderInfo(
                    location_type="paragraph",
                    paragraph_index=p_idx,
                    run_index=run_idx,
                    placeholder_pattern=match.group(),
                    char_offset=match.start(),
                    context_before=text[: match.start()].strip(),
                    context_after=text[match.end() :].strip(),
                    full_paragraph_text=text,
                    surrounding_context=_get_surrounding_context(container.paragraphs, p_idx, text),
                    section_index=section_index,
                    header_footer_type=header_footer_type,
                )
            )

    # Scan Tables
    for t_idx, table in enumerate(container.tables):
        # Extract headers for this table if possible
        headers = []
        for row in table.rows:
            texts = [c.text.strip() for c in row.cells]
            if any(t and not ANY_BLANK_RE.fullmatch(t) for t in texts):
                headers = texts
                break

        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                col_header = headers[c_idx] if c_idx < len(headers) else ""

                # Scan paragraphs inside cell
                for p_idx, para in enumerate(cell.paragraphs):
                    text = para.text
                    # Checkboxes in table cells - WE NEED TO ADD THIS, WAS MISSING
                    for match in CHECKBOX_RE.finditer(text):
                        run_idx = _find_run_for_position(para, match.start())
                        after = text[match.end() :].strip()
                        label = after.split("\n")[0].strip() if after else ""

                        placeholders.append(
                            PlaceholderInfo(
                                location_type="table",
                                table_index=t_idx,
                                row_index=r_idx,
                                cell_index=c_idx,
                                cell_paragraph_index=p_idx,
                                column_header=col_header,
                                paragraph_index=-1,
                                run_index=run_idx,
                                placeholder_pattern=match.group(),
                                char_offset=match.start(),
                                context_before=text[: match.start()].strip(),
                                context_after=text[match.end() :].strip(),
                                full_paragraph_text=text,
                                is_checkbox=True,
                                checkbox_label=label,
                                surrounding_context=_get_surrounding_context(cell.paragraphs, p_idx, text),
                                section_index=section_index,
                                header_footer_type=header_footer_type,
                            )
                        )

                    # Text Blanks in table cells
                    for match in ANY_BLANK_RE.finditer(text):
                        placeholders.append(
                            PlaceholderInfo(
                                location_type="table",
                                table_index=t_idx,
                                row_index=r_idx,
                                cell_index=c_idx,
                                cell_paragraph_index=p_idx,
                                column_header=col_header,
                                paragraph_index=-1,
                                run_index=_find_run_for_position(para, match.start()),
                                placeholder_pattern=match.group(),
                                char_offset=match.start(),
                                context_before=text[: match.start()].strip(),
                                context_after=text[match.end() :].strip(),
                                full_paragraph_text=text,
                                surrounding_context=_get_surrounding_context(cell.paragraphs, p_idx, text),
                                section_index=section_index,
                                header_footer_type=header_footer_type,
                            )
                        )

                # Nested tables
                if cell.tables:
                    _scan_container(
                        cell, placeholders, section_index=section_index, header_footer_type=header_footer_type
                    )


def paragraph_runs(paragraph: Paragraph) -> list[Run]:
    """Text runs in display order, including runs inside hyperlinks."""
    runs = []
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            runs.extend(item.runs)
        else:
            runs.append(item)
    return runs


def _find_run_for_position(paragraph, char_pos: int) -> int:
    """Find which run contains the character at the given position."""
    current_pos = 0
    for r_idx, run in enumerate(paragraph_runs(paragraph)):
        run_len = len(run.text)
        if current_pos + run_len > char_pos:
            return r_idx
        current_pos += run_len
    return max(0, len(paragraph_runs(paragraph)) - 1)


def _find_cell_run(cell, char_pos: int) -> tuple[int, int]:
    """Find the paragraph index and run index within a table cell for a character position."""
    current_pos = 0
    for p_idx, para in enumerate(cell.paragraphs):
        for r_idx, run in enumerate(paragraph_runs(para)):
            run_len = len(run.text)
            if current_pos + run_len > char_pos:
                return p_idx, r_idx
            current_pos += run_len
        current_pos += 1
    return 0, 0


def _get_table_context(doc: Document, table_index: int) -> str:
    """Get surrounding paragraph text for a table to understand its context."""
    from docx.oxml.ns import qn

    body = doc.element.body
    tables_seen = 0
    context_parts = []

    for elem in body:
        if elem.tag == qn("w:tbl"):
            if tables_seen == table_index:
                break
            tables_seen += 1
        elif elem.tag == qn("w:p"):
            full_text = ""
            for r in elem.findall(f".//{qn('w:t')}"):
                full_text += r.text or ""
            if full_text.strip():
                context_parts.append(full_text.strip())
                if len(context_parts) > 3:
                    context_parts.pop(0)

    return " | ".join(context_parts)


def extract_data_tables(doc: Document) -> list[TableInfo]:
    """Identify tables that need to be filled with array data (e.g., services table).

    Handles multi-row headers by selecting the header row with the most
    unique non-empty cells (i.e., the per-column header row, not merged rows).
    """
    tables = []

    for t_idx, table in enumerate(doc.tables):
        if len(table.rows) < 3:
            continue

        header_candidates = []  # list of (row_index, texts)
        data_rows = []
        found_data = False

        for r_idx, row in enumerate(table.rows):
            texts = [cell.text.strip() for cell in row.cells]
            all_empty = all(not t for t in texts)
            has_content = any(t and not ANY_BLANK_RE.fullmatch(t) for t in texts)

            if has_content and not found_data:
                # Skip rows where all cells are identical (merged row spanning all columns)
                if len(set(texts)) == 1:
                    continue
                header_candidates.append((r_idx, texts))
            elif all_empty and header_candidates:
                found_data = True
                data_rows.append(r_idx)
            elif all_empty and found_data:
                data_rows.append(r_idx)

        if header_candidates and data_rows:
            # Pick the header row with the most unique non-empty cells
            # This prefers the detailed per-column row over merged header rows
            best_headers = header_candidates[0][1]
            best_score = 0
            for _, texts in header_candidates:
                non_empty = [t for t in texts if t]
                unique_non_empty = set(non_empty)
                score = len(unique_non_empty)
                if score > best_score:
                    best_score = score
                    best_headers = texts

            context = _get_table_context(doc, t_idx)
            tables.append(
                TableInfo(
                    table_index=t_idx,
                    headers=best_headers,
                    data_row_indices=data_rows,
                    context=context,
                )
            )

    return tables


def get_all_placeholders(doc: Document) -> tuple[Document, list[PlaceholderInfo], list[TableInfo]]:
    """Get all placeholders using the new recursive strategy."""
    return extract_placeholders_recursive(doc)
