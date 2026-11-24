import json
import re
from copy import deepcopy

from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_analyzer import PlaceholderInfo, TableInfo, paragraph_runs
from docx_analyzer import placeholder_id as _placeholder_id

# Regex for finding placeholders
CHECKBOX_RE = re.compile(r"\|_\|")
ANY_BLANK_RE = re.compile(r"_+")


def fill_document(
    doc: Document,
    placeholders: list[PlaceholderInfo],
    mapping: dict,
    json_data: dict,
    data_tables: list[TableInfo] | None = None,
    table_key_mapping: dict[int, str] | None = None,
    column_mappings: dict[int, dict[int, str]] | None = None,
) -> tuple[int, list[dict], set[str], list[dict]]:
    """
    Fill the DOCX document in-place using the mapping (which now contains EXACT values).

    Args:
        table_key_mapping: LLM-provided mapping of table_index -> json_key
        column_mappings: LLM-provided mapping of table_index -> {col_index -> dict_key}

    Returns:
        tuple(filled_count, table_stats, table_used_keys, fill_log)
    """
    filled_count = 0
    fill_log = []

    if data_tables is None:
        data_tables = []

    # 1. Fill text placeholders and checkboxes in paragraphs
    count, p_logs = _fill_paragraph_placeholders(doc, placeholders, mapping)
    filled_count += count
    fill_log.extend(p_logs)

    # 2. Fill table cell placeholders
    count, t_logs = _fill_table_placeholders(doc, placeholders, mapping)
    filled_count += count
    fill_log.extend(t_logs)

    # 3. Fill data tables with array data
    table_stats, table_used_keys, d_logs = _fill_data_tables(
        doc, data_tables, json_data, table_key_mapping, column_mappings
    )
    fill_log.extend(d_logs)

    print(f"[Fill] Document filled successfully. {filled_count} placeholders replaced.")
    return filled_count, table_stats, table_used_keys, fill_log


def _get_container(doc: Document, section_index: int, header_footer_type: str):
    """Resolve the container (body, header, footer) based on location info."""
    if section_index == -1:
        return doc

    if section_index >= len(doc.sections):
        return None

    section = doc.sections[section_index]

    if header_footer_type == "body":
        return doc  # Paragraphs are in doc.paragraphs, but section-specific body isn't easily isolated without iterating all.
        # However, our analyzer stores absolute paragraph indices for body,
        # so we can just use doc.paragraphs[p_index] if we are careful.
        # BUT wait: doc.paragraphs contains ALL body paragraphs.
        # Yes, p.paragraph_index is index in doc.paragraphs.
        return doc

    # Header/Footer resolution
    try:
        if header_footer_type == "header_primary":
            return section.header
        elif header_footer_type == "header_first":
            return section.first_page_header
        elif header_footer_type == "header_even":
            return section.even_page_header
        elif header_footer_type == "footer_primary":
            return section.footer
        elif header_footer_type == "footer_first":
            return section.first_page_footer
        elif header_footer_type == "footer_even":
            return section.even_page_footer
    except Exception:
        return None

    return None


def _fill_paragraph_placeholders(
    doc: Document, placeholders: list[PlaceholderInfo], mapping: dict
) -> tuple[int, list[dict]]:
    """Fill text blanks and checkboxes in paragraphs."""

    # Group placeholders by (section_index, header_footer_type, paragraph_index)
    # This avoids resolving the container 1000 times
    grouped: dict[tuple, list[PlaceholderInfo]] = {}

    for p in placeholders:
        if p.location_type != "paragraph":
            continue
        key = (p.section_index, p.header_footer_type, p.paragraph_index)
        grouped.setdefault(key, []).append(p)

    # Pre-resolve paragraph objects to handle index shifting when new paragraphs are inserted
    # (e.g. from multiline values). We must grab references BEFORE modification.
    para_map = {}
    for s_idx, hf_type, p_idx in grouped.keys():
        container = _get_container(doc, s_idx, hf_type)
        if not container:
            continue
        if p_idx >= len(container.paragraphs):
            continue
        para_map[(s_idx, hf_type, p_idx)] = container.paragraphs[p_idx]

    filled_count = 0
    checkbox_count = 0
    logs = []

    # Sort keys to process in order (though object references make order less critical for correctness)
    for key in sorted(grouped.keys()):
        p_list = grouped[key]
        para = para_map.get(key)

        if not para:
            continue

        for placeholder in sorted(p_list, key=lambda field: field.char_offset, reverse=True):
            pid = _placeholder_id(placeholder)
            val_obj = mapping.get(pid)

            # If value is None, skip filling
            # Now mapping values are dicts {value, used_keys}
            if not val_obj or val_obj.get("value") is None:
                continue

            value = str(val_obj["value"])
            used_keys = val_obj.get("used_keys", [])

            if placeholder.is_checkbox:
                # Handle checkbox: value should be "CHECKED" or "UNCHECKED"
                _fill_checkbox(para, placeholder, value)
                checkbox_count += 1
            else:
                # Handle text blank: replace the blank pattern with the value
                if _fill_text_blank(para, placeholder, value):
                    filled_count += 1

            logs.append(
                {
                    "type": "placeholder",
                    "id": pid,
                    "value": value,
                    "source": used_keys,
                    "context": placeholder.context_before + placeholder.placeholder_pattern + placeholder.context_after,
                }
            )

    print(f"[Fill] Filled {filled_count} text blanks and {checkbox_count} checkboxes in paragraphs")
    return filled_count + checkbox_count, logs


def _split_paragraph_at_newlines(paragraph, run, value, text_before="", text_after=""):
    """
    When a replacement value contains \n, split it into multiple DOCX paragraphs.
    Preserves paragraph formatting (<w:pPr>) and run formatting (<w:rPr>).
    Moves subsequent runs to the last new paragraph to ensure text after the newline
    but in different runs is correctly positioned.

    Args:
        paragraph: the paragraph containing the run
        run: the run where the replacement starts
        value: the replacement text (contains \n)
        text_before: text in the run before the replaced pattern
        text_after: text in the run after the replaced pattern
    """
    lines = value.split("\n")

    # 1. First line stays in the existing run
    run.text = text_before + lines[0]

    # 2. Identify subsequent runs to move
    # We need to find the index of 'run' in 'paragraph.runs'
    # Use XML element identity since Run objects might be regenerated
    run_index = -1
    all_runs = paragraph.runs
    for i, r in enumerate(all_runs):
        if r._r == run._r:
            run_index = i
            break

    runs_to_move = []
    if run_index != -1 and run_index < len(all_runs) - 1:
        # These are the runs AFTER the split point that need to move to the last paragraph
        runs_to_move = all_runs[run_index + 1 :]

    # Get formatting to clone
    p_element = paragraph._p
    pPr = p_element.find(qn("w:pPr"))
    rPr = run._r.find(qn("w:rPr"))

    current_element = p_element

    for i, line in enumerate(lines[1:], 1):
        new_p = OxmlElement("w:p")
        if pPr is not None:
            new_pPr = deepcopy(pPr)
            # Remove w:pId if present to avoid duplicate IDs
            pId = new_pPr.find(qn("w:pId"))
            if pId is not None:
                new_pPr.remove(pId)
            new_p.append(new_pPr)

        new_r = OxmlElement("w:r")
        if rPr is not None:
            new_r.append(deepcopy(rPr))

        # Last line gets text_after appended
        is_last_line = i == len(lines) - 1
        line_text = line + text_after if is_last_line else line

        new_t = OxmlElement("w:t")
        new_t.text = line_text
        if line_text and (line_text[0] == " " or line_text[-1] == " "):
            new_t.set(qn("xml:space"), "preserve")

        new_r.append(new_t)
        new_p.append(new_r)

        # If this is the last line, append the moved runs
        if is_last_line and runs_to_move:
            for r_move in runs_to_move:
                # Appending the XML element to the new paragraph automatically
                # removes it from the old one in lxml
                new_p.append(r_move._r)

        current_element.addnext(new_p)
        current_element = new_p


def _fill_text_blank(paragraph, placeholder: PlaceholderInfo, value: str) -> bool:
    """
    Replace the exact blank pattern (e.g. "___________") with the value.
    Handling runs is tricky because the blank might be split across runs.
    If value contains line breaks, splits into multiple paragraphs.
    """
    pattern = placeholder.placeholder_pattern
    if placeholder.char_offset >= 0:
        return _replace_pattern_across_runs(paragraph, pattern, value, placeholder.char_offset)
    value = value.replace("\r\n", "\n").replace("\r", "\n")

    # 1. Try simple replacement (if blank is entirely in one run)
    if placeholder.run_index >= 0 and placeholder.run_index < len(paragraph.runs):
        run = paragraph.runs[placeholder.run_index]
        if pattern in run.text:
            if "\n" in value:
                idx = run.text.find(pattern)
                text_before = run.text[:idx]
                text_after = run.text[idx + len(pattern) :]
                _split_paragraph_at_newlines(paragraph, run, value, text_before, text_after)
            else:
                run.text = run.text.replace(pattern, value, 1)
            return True

    # 2. If not found in single run, it might be split across runs.
    # Fallback: Searching across runs
    return _replace_pattern_across_runs(paragraph, pattern, value)


def _replace_pattern_across_runs(paragraph, pattern: str, value: str, start_idx: int | None = None) -> bool:
    """
    Find identifiers split across runs and replace them.
    This is a naive implementation: it finds the start and end run indices
    and replaces the text.
    """
    runs = paragraph_runs(paragraph)
    full_text = "".join(r.text for r in runs)
    if pattern not in full_text:
        return False

    if start_idx is None:
        start_idx = full_text.find(pattern)
    if full_text[start_idx : start_idx + len(pattern)] != pattern:
        return False
    end_idx = start_idx + len(pattern)

    # Find which runs cover this range
    current_pos = 0
    start_run = -1
    end_run = -1

    local_start = 0
    local_end = 0

    for i, run in enumerate(runs):
        run_len = len(run.text)
        run_end_pos = current_pos + run_len

        if start_run == -1 and start_idx < run_end_pos:
            start_run = i
            local_start = start_idx - current_pos

        if start_run != -1 and end_idx <= run_end_pos:
            end_run = i
            local_end = end_idx - current_pos
            break

        current_pos += run_len

    if start_run != -1 and end_run != -1:
        value_norm = value.replace("\r\n", "\n").replace("\r", "\n")
        has_newlines = "\n" in value_norm

        if start_run == end_run:
            run = runs[start_run]
            if has_newlines:
                text_before = run.text[:local_start]
                text_after = run.text[local_end:]
                _split_paragraph_at_newlines(paragraph, run, value_norm, text_before, text_after)
            else:
                run.text = run.text[:local_start] + value + run.text[local_end:]
        else:
            if has_newlines:
                text_before = runs[start_run].text[:local_start]
                text_after = runs[end_run].text[local_end:]
                # Clear middle and end runs
                for i in range(start_run + 1, end_run + 1):
                    runs[i].text = ""
                _split_paragraph_at_newlines(paragraph, runs[start_run], value_norm, text_before, text_after)
            else:
                runs[start_run].text = runs[start_run].text[:local_start] + value
                for i in range(start_run + 1, end_run):
                    runs[i].text = ""
                runs[end_run].text = runs[end_run].text[local_end:]

    return True


def _fill_checkbox(paragraph, placeholder: PlaceholderInfo, value: str) -> None:
    """
    Fill a checkbox by replacing |_| with |X| or leaving |_| based on the value.

    Logic: value from LLM should be "CHECKED" or "UNCHECKED".
    """
    if value.strip().upper() == "CHECKED":
        offset = placeholder.char_offset if placeholder.char_offset >= 0 else None
        _replace_pattern_across_runs(paragraph, "|_|", "|X|", offset)


def _replace_checkbox_xml(paragraph) -> None:
    """Replace checkbox at XML level if runs don't directly contain |_|."""
    xml_str = paragraph._p.xml
    if "|_|" in xml_str:
        # The checkbox might be split across runs
        # Note: altering XML directly is complex, for now we rely on run replacement
        pass


def _fill_table_placeholders(
    doc: Document, placeholders: list[PlaceholderInfo], mapping: dict
) -> tuple[int, list[dict]]:
    """Fill blanks in table cells."""
    filled_count = 0
    logs = []

    ordered = sorted(
        placeholders,
        key=lambda field: (
            field.section_index,
            field.header_footer_type,
            field.table_index,
            field.row_index,
            field.cell_index,
            field.cell_paragraph_index,
            field.char_offset,
        ),
        reverse=True,
    )
    for placeholder in ordered:
        if placeholder.location_type != "table":
            continue

        pid = _placeholder_id(placeholder)
        val_obj = mapping.get(pid)

        if not val_obj or val_obj.get("value") is None:
            continue

        # Ensure string
        value = str(val_obj["value"])
        used_keys = val_obj.get("used_keys", [])

        container = _get_container(doc, placeholder.section_index, placeholder.header_footer_type)
        if not container:
            continue

        if placeholder.table_index >= len(container.tables):
            continue
        table = container.tables[placeholder.table_index]

        if placeholder.row_index >= len(table.rows):
            continue
        row = table.rows[placeholder.row_index]
        if placeholder.cell_index >= len(row.cells):
            continue
        cell = row.cells[placeholder.cell_index]

        if not 0 <= placeholder.cell_paragraph_index < len(cell.paragraphs):
            continue
        para = cell.paragraphs[placeholder.cell_paragraph_index]
        filled = False
        if placeholder.is_checkbox:
            if value in {"CHECKED", "UNCHECKED"}:
                _fill_checkbox(para, placeholder, value)
                filled = True
        else:
            filled = _fill_text_blank(para, placeholder, value)
        if filled:
            filled_count += 1

        if filled:
            logs.append(
                {
                    "type": "placeholder_table",
                    "id": pid,
                    "value": value,
                    "source": used_keys,
                    "table_idx": placeholder.table_index,
                    "row": placeholder.row_index,
                    "col": placeholder.cell_index,
                }
            )

    print(f"[Fill] Filled {filled_count} table cell placeholders")
    return filled_count, logs


def _fill_data_tables(
    doc: Document,
    data_tables: list[TableInfo],
    json_data: dict,
    table_key_mapping: dict[int, str] | None = None,
    column_mappings: dict[int, dict[int, str]] | None = None,
) -> tuple[list[dict], set[str], list[dict]]:
    """
    Fill data tables (like the services table) with array data from JSON.

    Args:
        table_key_mapping: LLM-provided mapping of table_index -> json_key
        column_mappings: LLM-provided mapping of table_index -> {col_index -> dict_key}

    Returns:
        tuple(stats_list, used_keys_set, detailed_logs)
    """
    stats: list[dict] = []
    used_keys: set[str] = set()
    logs: list[dict] = []

    # Find JSON keys that contain array data
    array_keys = {}
    for key, value in json_data.items():
        if isinstance(value, list):
            array_keys[key] = value
        elif isinstance(value, str) and value.strip().startswith("["):
            try:
                parsed = json.loads(value)
                if isinstance(parsed, list):
                    array_keys[key] = parsed
            except json.JSONDecodeError:
                continue

    if not array_keys:
        return stats, used_keys, logs

    if table_key_mapping is None:
        table_key_mapping = {}
    if column_mappings is None:
        column_mappings = {}

    for table_info in data_tables:
        if table_info.table_index >= len(doc.tables):
            continue
        table = doc.tables[table_info.table_index]

        # Use LLM-provided mapping, fall back to heuristic if not available
        matched_key = table_key_mapping.get(table_info.table_index)
        if matched_key is None or matched_key not in array_keys:
            matched_key = _match_table_to_json_key(table_info, array_keys)
        if matched_key is None:
            continue

        array_data = array_keys[matched_key]
        rows_filled = 0
        used_keys.add(matched_key)

        # Get column mapping for this table
        col_mapping = column_mappings.get(table_info.table_index, {})

        print(f"[Fill] Filling data table {table_info.table_index} with '{matched_key}' ({len(array_data)} rows)")

        # Snapshot before any values are written: overflow rows must not inherit data.
        template_row = (
            deepcopy(table.rows[table_info.data_row_indices[-1]]._tr) if table_info.data_row_indices else None
        )
        # Fill data rows
        for data_idx, data_row in enumerate(array_data):
            # If we run out of pre-allocated rows, add new ones by cloning the last valid data row
            if data_idx >= len(table_info.data_row_indices):
                if not table_info.data_row_indices:
                    break
                row_idx = table_info.data_row_indices[-1] + data_idx - len(table_info.data_row_indices) + 1
                _add_row_to_table(table, template_row, row_idx)
            else:
                row_idx = table_info.data_row_indices[data_idx]

            if row_idx >= len(table.rows):
                continue

            table_row = table.rows[row_idx]

            if isinstance(data_row, dict):
                if col_mapping:
                    # Use LLM-provided column mapping
                    row_fills = _fill_table_row_from_mapping(table_row, col_mapping, data_row)
                else:
                    # Fallback to heuristic
                    row_fills = _fill_table_row_from_dict(table_row, table_info.headers, data_row)
                # Log detailed fills
                for col_idx, val, key in row_fills:
                    logs.append(
                        {
                            "type": "table_cell",
                            "table_idx": table_info.table_index,
                            "row": row_idx,
                            "col": col_idx,
                            "value": val,
                            "source": f"{matched_key}[{data_idx}].{key}",
                        }
                    )
            elif isinstance(data_row, str):
                # Simple string — fill first empty cell
                for col_idx, cell in enumerate(table_row.cells):
                    if not cell.text.strip():
                        _set_cell_text(cell, data_row)
                        logs.append(
                            {
                                "type": "table_cell",
                                "table_idx": table_info.table_index,
                                "row": row_idx,
                                "col": col_idx,
                                "value": data_row,
                                "source": f"{matched_key}[{data_idx}]",
                            }
                        )
                        break
            rows_filled += 1

        stats.append(
            {
                "table_index": table_info.table_index,
                "json_key": matched_key,
                "rows_filled": rows_filled,
                "total_rows": len(array_data),
            }
        )

    return stats, used_keys, logs


def _fill_table_row_from_mapping(table_row, column_mapping: dict[int, str], data: dict) -> list[tuple]:
    """
    Fill a table row using LLM-provided column mapping.
    column_mapping: {col_index -> dict_key}
    Returns list of (col_index, value, key_used) for logging.
    """
    fills = []
    for col_idx, cell in enumerate(table_row.cells):
        if col_idx in column_mapping:
            dk = column_mapping[col_idx]
            raw = data.get(dk)
            val = "" if raw is None else str(raw)
            _set_cell_text(cell, val)
            fills.append((col_idx, val, dk))
    return fills


def _match_table_to_json_key(table_info: TableInfo, array_keys: dict) -> str | None:
    """Match a data table to the right JSON array key using headers, context, and data structure.

    Uses diacritics-normalized matching and gives strong bonuses when dict keys
    in the JSON data structurally match the table column headers.
    """
    context_norm = _normalize_text(table_info.context)
    headers_norm = " ".join(_normalize_text(h.replace("\n", " ")) for h in table_info.headers)

    best_match = None
    best_score = 0

    for key, data in array_keys.items():
        score = 0
        key_norm = _normalize_text(key.replace("-", " ").replace("_", " "))

        # Check if key words appear in context or headers
        key_words = key_norm.split()
        for word in key_words:
            if len(word) > 3:  # skip short words
                if word in context_norm:
                    score += 2
                if word in headers_norm:
                    score += 3

        # Strong bonus: check if dict keys in data match column headers
        # This is the most reliable signal — structural match between JSON and table
        if data and isinstance(data[0], dict):
            data_keys_norm = {dk: _normalize_text(dk.replace("_", " ")) for dk in data[0].keys()}
            for header in table_info.headers:
                header_norm = _normalize_text(header.replace("\n", " "))
                if not header_norm:
                    continue
                header_words = set(header_norm.split())

                for _, dk_norm in data_keys_norm.items():
                    dk_words = set(dk_norm.split())
                    # Word overlap
                    overlap = dk_words & header_words
                    if overlap:
                        score += len(overlap) * 3
                    # Prefix match (for Romanian morphology: ofertant/ofertantului)
                    for dw in dk_words:
                        for hw in header_words:
                            if dw == hw:
                                continue
                            if len(dw) >= 4 and len(hw) >= 4 and (dw.startswith(hw[:4]) or hw.startswith(dw[:4])):
                                score += 2

        if score > best_score:
            best_score = score
            best_match = key

    return best_match if best_score > 0 else None


def _normalize_text(text: str) -> str:
    """Normalize text by removing diacritics, punctuation and lowering case for fuzzy matching."""
    import unicodedata

    # Decompose unicode characters (e.g., Ţ -> T + cedilla)
    nfkd = unicodedata.normalize("NFKD", text.lower())
    # Remove combining characters (accents, cedillas, etc.)
    ascii_text = "".join(c for c in nfkd if not unicodedata.combining(c))
    # Remove punctuation (parentheses, slashes, commas, etc.)
    return re.sub(r"[^\w\s]", "", ascii_text).strip()


def _fill_table_row_from_dict(table_row, headers: list[str], data: dict) -> list[tuple]:
    """
    Fill a table row using a dict, mapping dict keys to column headers.
    Returns list of (col_index, value, key_used) for logging.

    Uses diacritics-normalized fuzzy matching and greedy 1:1 assignment
    to prevent the same key being mapped to multiple columns.
    """
    # Build normalized versions
    header_norms = []
    for header in headers:
        clean = header.lower().strip().replace("\n", " ")
        header_norms.append(_normalize_text(clean))

    dk_norms = {}
    for dk in data.keys():
        dk_norms[dk] = _normalize_text(dk.replace("_", " "))

    # Score all (header_idx, data_key) pairs
    scored_pairs = []
    for h_idx in range(len(headers)):
        h_norm = header_norms[h_idx]
        if not h_norm:
            continue

        h_words = set(h_norm.split())

        for dk, dk_norm in dk_norms.items():
            dk_words = set(dk_norm.split())
            score = 0

            # Exact word overlap
            overlap = dk_words & h_words
            score += len(overlap) * 3

            # Partial word matching (prefix of >= 4 chars)
            for dw in dk_words:
                for hw in h_words:
                    if dw == hw:
                        continue  # already counted
                    # Check if one is a prefix of the other (min 4 chars)
                    min_len = min(len(dw), len(hw))
                    if min_len >= 4 and (dw.startswith(hw[:4]) or hw.startswith(dw[:4])):
                        score += 2

            # Full substring containment
            if dk_norm in h_norm or h_norm in dk_norm:
                score += 2

            if score > 0:
                scored_pairs.append((score, h_idx, dk))

    # Greedy 1:1 assignment: highest score first, each key and header used at most once
    scored_pairs.sort(key=lambda x: -x[0])
    header_to_key = {}
    used_keys = set()

    for _score, h_idx, dk in scored_pairs:
        header = headers[h_idx]
        if header in header_to_key or dk in used_keys:
            continue
        header_to_key[header] = dk
        used_keys.add(dk)

    # Fill cells
    fills = []
    for c_idx, cell in enumerate(table_row.cells):
        if c_idx < len(headers):
            header = headers[c_idx]
            if header in header_to_key:
                raw = data.get(header_to_key[header])
                val = "" if raw is None else str(raw)
                _set_cell_text(cell, val)
                fills.append((c_idx, val, header_to_key[header]))
    return fills


def _set_cell_text(cell, text: str) -> None:
    """
    Set cell text preserving existing formatting.
    Uses the first paragraph's formatting as template.
    If text contains line breaks, creates multiple paragraphs.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    if cell.paragraphs:
        para = cell.paragraphs[0]
        if "\n" in text:
            # Multi-line: split into paragraphs preserving formatting
            if para.runs:
                run = para.runs[0]
                for r in para.runs[1:]:
                    r.text = ""
            else:
                run = para.add_run("")
            _split_paragraph_at_newlines(para, run, text)
        else:
            if para.runs:
                # Preserve existing run formatting
                para.runs[0].text = text
                # Clear any additional runs
                for run in para.runs[1:]:
                    run.text = ""
            else:
                # No runs exist — add one with paragraph formatting
                run = para.add_run(text)
    else:
        cell.text = text


def _add_row_to_table(table, template_row, insert_index: int):
    """
    Add a new row to the table by cloning the template row.
    Preserves cell formatting (borders, shading, etc.) by copying the XML element.
    """
    new_tr = deepcopy(template_row)
    # Keep totals, notes, and signature rows after the expandable data region.
    table.rows[insert_index - 1]._tr.addnext(new_tr)
    return table.rows[insert_index]
