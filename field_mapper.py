"""
Field Mapper — Uses LLM to determine the EXACT value to fill for each DOCX placeholder from JSON.

Sends DOCX placeholder contexts and JSON data to the configured OpenAI model,
which returns the best string value for each placeholder. The mapping is
cached to .cache/mapping.json for deterministic subsequent runs.
"""

import hashlib
import json
import os
from dataclasses import asdict
from pathlib import Path

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI

from docx_analyzer import PlaceholderInfo, TableInfo
from docx_analyzer import placeholder_id as _placeholder_id

CACHE_DIR = Path(".cache")
CACHE_FILE = CACHE_DIR / "mapping.json"
TABLE_CACHE_FILE = CACHE_DIR / "table_mapping.json"


def _model_name() -> str:
    return os.environ.get("OPENAI_MODEL", "gpt-5.1")


def _read_cache(path: Path) -> dict:
    """Treat malformed or unreadable cache files as a miss."""
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError):
        return {}


def _response_items(content: object) -> list[dict]:
    """Parse model output without trusting its JSON shape."""
    text = str(content).strip()
    if text.startswith("```"):
        text = text.partition("\n")[2]
        if text.endswith("```"):
            text = text[:-3].strip()
    try:
        value = json.loads(text)
    except ValueError:
        return []
    return [item for item in value if isinstance(item, dict)] if isinstance(value, list) else []


def _compute_inputs_hash(placeholders: list[PlaceholderInfo], json_data: dict) -> str:
    """Compute a hash of the inputs to detect changes that invalidate the cache."""
    # We now depend on the ACTUAL DATA in json_data, not just keys
    data = json.dumps(
        {
            "placeholders": [asdict(p) for p in placeholders],
            "json_data": json_data,
            "version": "v3_context_and_model",
            "model": _model_name(),
        },
        sort_keys=True,
    )
    return hashlib.sha256(data.encode()).hexdigest()[:16]


def load_cached_mapping(placeholders: list[PlaceholderInfo], json_data: dict) -> dict[str, dict] | None:
    """Load mapping from cache if it exists and inputs haven't changed."""
    cache = _read_cache(CACHE_FILE)
    if cache.get("inputs_hash") != _compute_inputs_hash(placeholders, json_data):
        return None
    mapping = cache.get("mapping")
    if not isinstance(mapping, dict):
        return None
    fields = {_placeholder_id(field): field for field in placeholders}
    for pid, decision in mapping.items():
        if pid not in fields or not isinstance(decision, dict):
            return None
        value, keys = decision.get("value"), decision.get("used_keys")
        if value is not None and (
            not isinstance(value, str) or (fields[pid].is_checkbox and value not in {"CHECKED", "UNCHECKED"})
        ):
            return None
        if not isinstance(keys, list) or not all(isinstance(key, str) and key in json_data for key in keys):
            return None
    return mapping


def save_mapping_to_cache(mapping: dict, placeholders: list[PlaceholderInfo], json_data: dict) -> None:
    """Save the mapping to cache for deterministic subsequent runs."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)

    cache = {
        "inputs_hash": _compute_inputs_hash(placeholders, json_data),
        "mapping": mapping,
    }
    CACHE_FILE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"[Cache] Saved mapping to {CACHE_FILE}")


def _build_placeholder_descriptions(placeholders: list[PlaceholderInfo]) -> list[dict]:
    """Build human-readable descriptions of placeholders for the LLM prompt."""
    descriptions = []
    for p in placeholders:
        pid = _placeholder_id(p)

        if p.is_checkbox:
            desc = {
                "id": pid,
                "type": "checkbox",
                "label": p.checkbox_label,
                "full_context": p.full_paragraph_text[:300],
                "nearby_paragraphs": p.surrounding_context[:300],
            }
        elif p.location_type == "table":
            desc = {
                "id": pid,
                "type": "table_cell",
                "column_header": p.column_header,
                "context_before": p.context_before[:150],
                "context_after": p.context_after[:150],
                "full_context": p.full_paragraph_text[:300],
            }
        else:
            desc = {
                "id": pid,
                "type": "text_blank",
                "context_before": p.context_before[:200],
                "context_after": p.context_after[:200],
                "full_context": p.full_paragraph_text[:800],
                "nearby_paragraphs": p.surrounding_context[:800],
            }
        descriptions.append(desc)

    return descriptions


def map_fields_with_llm(placeholders: list[PlaceholderInfo], json_keys: list[str], json_data: dict) -> dict[str, dict]:
    """
    Use LLM to determine the EXACT value to fill for each DOCX placeholder from JSON.

    Returns:
        - mapping: dict placeholder_id -> { "value": str|None, "used_keys": list[str] }
    """
    # Check cache first
    cached = load_cached_mapping(placeholders, json_data)
    if cached is not None:
        mapping = cached
        _print_mapping_report(mapping)
        return mapping

    print(f"[LLM] Mapping {len(placeholders)} placeholders using full JSON data...")

    # Build descriptions
    descriptions = _build_placeholder_descriptions(placeholders)

    # Create LLM
    llm = ChatOpenAI(
        model=_model_name(),
        temperature=0,
    )

    # Split into batches to avoid token limits (5 placeholders per batch)
    BATCH_SIZE = 10
    all_mappings: dict[str, dict] = {}

    # We pass the full JSON object now, not just keys
    json_context = json.dumps(json_data, indent=2, ensure_ascii=False)

    for batch_start in range(0, len(descriptions), BATCH_SIZE):
        batch = descriptions[batch_start : batch_start + BATCH_SIZE]
        batch_num = batch_start // BATCH_SIZE + 1
        total_batches = (len(descriptions) + BATCH_SIZE - 1) // BATCH_SIZE

        print(f"  [LLM] Processing batch {batch_num}/{total_batches} ({len(batch)} placeholders)...")

        system_prompt = """You are an expert at filling procurement documents (DOCX) from JSON data.

Given a list of DOCX placeholder positions (with context) and a JSON data object, determine the EXACT string value to fill in each placeholder.

RULES for Generating Values:
1. **Strict Fidelity**: Use data ONLY from the provided JSON. Do not invent information.
2. **Smart Extraction**: 
   - If a placeholder requires a combination of fields (e.g. "Name, Address"), combine them intelligently from the JSON.
   - Use semantic reasoning to find the most relevant data.
   - If the JSON field is a sentence (e.g., 'ID Series RX No 12345') and the placeholder needs only the number, EXTRACT '12345'. Do not return the whole string.
   - If a placeholder needs a part of a compound value (e.g., just the 'Day' from a Date '2025-10-08'), extract that part.
   - Ignore formatting artifacts in the placeholder context (like '......' or '_____'). Focus on the semantic meaning of the sentence around it.
   - If multiple similar placeholders appear in a sequence (e.g., a list of names or items), and the JSON contains corresponding list of values, map them sequentially 1-to-1.
   - Treat all visual placeholders (dots ......, underscores ______, lines) purely as insertion points. Their length or character style does NOT indicate the length of the data to be inserted.
   - If an exact label match is missing, use the most semantically equivalent (if exists) field available in the JSON
3. **Checkboxes**:
   - If functionality `type` is "checkbox", return exactly "CHECKED" or "UNCHECKED".
   - Decide based on the JSON data and the checkbox label.
   - Example: specific checkbox for "IMM" should be checked if JSON says company type is IMM.
4. **Empty/Null**:
   - If no relevant data exists in JSON for a placeholder, return null (JSON null).
   - If the JSON value is null/empty, return null.
5. **Preserve Formatting**:
   - If the JSON value contains newlines (\n), YOU MUST PRESERVE THEM in the output value. Do not replace them with spaces.

Return a JSON array of objects, each with:
- "id": the placeholder id (exactly as given)
- "value": the exact string to put in the docx, "CHECKED"/"UNCHECKED" for checkboxes, or null.
- "used_json_keys": list of strings, names of the keys from the JSON object used to derive this value. If value is null, list is empty.
"""

        human_prompt = f"""Here represents the JSON Data source:
{json_context}

Here are the DOCX placeholders to fill:
{json.dumps(batch, indent=2, ensure_ascii=False)}

Return a JSON array mapping each placeholder id to its "value".
Output ONLY the JSON array, no markdown formatting.
"""

        response = llm.invoke(
            [
                SystemMessage(content=system_prompt),
                HumanMessage(content=human_prompt),
            ]
        )

        # Accept only requested IDs and known JSON keys; leave invalid entries unfilled.
        requested = {desc["id"] for desc in batch}
        checkbox_ids = {_placeholder_id(field) for field in placeholders if field.is_checkbox}
        for pid in requested:
            all_mappings[pid] = {"value": None, "used_keys": []}
        for item in _response_items(response.content):
            pid = item.get("id")
            value = item.get("value")
            keys = item.get("used_json_keys")
            if not isinstance(pid, str) or pid not in requested:
                continue
            if value is not None and (
                not isinstance(value, str) or (pid in checkbox_ids and value not in {"CHECKED", "UNCHECKED"})
            ):
                continue
            if not isinstance(keys, list) or not all(isinstance(key, str) and key in json_data for key in keys):
                continue
            all_mappings[pid] = {"value": value, "used_keys": keys}

    # Save to cache
    save_mapping_to_cache(all_mappings, placeholders, json_data)

    # Print detailed report
    _print_mapping_report(all_mappings)

    return all_mappings


def _print_mapping_report(mapping: dict) -> None:
    """Print detailed matching statistics to the terminal."""
    total = len(mapping)
    matched_placeholders = {k: v for k, v in mapping.items() if v.get("value") is not None}
    unmatched_placeholders = [k for k, v in mapping.items() if v.get("value") is None]

    # Collect all used keys
    all_used_keys = set()
    for v in mapping.values():
        all_used_keys.update(v.get("used_keys", []))

    # Header
    print()
    print("=" * 60)
    print("  FILLING REPORT")
    print("=" * 60)

    # DOCX placeholders stats
    print(f"\n📄 DOCX Placeholders:  {total} total")
    print(f"   ✅ To be filled:    {len(matched_placeholders)}")
    print(f"   ❌ Unmatched/Empty: {len(unmatched_placeholders)}")

    # Used keys stats
    print(f"   🔑 JSON Fields Used: {len(all_used_keys)}")

    if matched_placeholders:
        print("\n--- Examples of values to fill ---")
        # Show first 5 examples
        for _, (pid, val_obj) in enumerate(list(matched_placeholders.items())[:5]):
            print(f"   • {pid[-20:]}: {val_obj['value']!r} (Source: {val_obj['used_keys']})")

    print()
    print("=" * 60)


def _compute_table_inputs_hash(data_tables: list[TableInfo], array_keys: dict) -> str:
    """Compute a hash of table mapping inputs for caching."""
    data = json.dumps(
        {
            "tables": [asdict(table) for table in data_tables],
            "array_keys": array_keys,
            "version": "v2_data_and_model",
            "model": _model_name(),
        },
        sort_keys=True,
    )
    return hashlib.sha256(data.encode()).hexdigest()[:16]


def map_table_to_json_key_with_llm(
    data_tables: list[TableInfo],
    array_keys: dict,
) -> dict[int, str]:
    """
    Use LLM to match each DOCX data table to the correct JSON array key.

    Returns:
        dict mapping table_index -> json_key
    """
    if not data_tables or not array_keys:
        return {}

    cache = _read_cache(TABLE_CACHE_FILE)
    valid_indices = {table.table_index for table in data_tables}
    if cache.get("inputs_hash") == _compute_table_inputs_hash(data_tables, array_keys):
        cached = cache.get("table_key_mapping")
        if isinstance(cached, dict):
            try:
                result = {int(key): value for key, value in cached.items()}
                if all(
                    key in valid_indices and isinstance(value, str) and value in array_keys
                    for key, value in result.items()
                ):
                    return result
            except (TypeError, ValueError):
                pass

    print(f"[LLM] Matching {len(data_tables)} data tables to {len(array_keys)} JSON array keys...")

    # Build descriptions
    tables_desc = []
    for t in data_tables:
        headers_clean = [h.replace("\n", " ").strip() for h in t.headers if h.strip()]
        tables_desc.append(
            {
                "table_index": t.table_index,
                "context": t.context[:300],
                "column_headers": headers_clean,
                "num_data_rows": len(t.data_row_indices),
            }
        )

    keys_desc = {}
    for key, data in array_keys.items():
        if data and isinstance(data[0], dict):
            keys_desc[key] = {"type": "array_of_dicts", "dict_keys": list(data[0].keys()), "count": len(data)}
        elif data:
            keys_desc[key] = {"type": "array_of_strings", "sample": data[:3], "count": len(data)}
        else:
            keys_desc[key] = {"type": "empty_array", "count": 0}

    llm = ChatOpenAI(model=_model_name(), temperature=0)

    system_prompt = """You are an expert at matching Romanian procurement document tables to JSON data.

Given a list of DOCX table descriptions (with column headers and context) and a list of JSON array keys (with their data structure), match each table to the correct JSON key.

RULES:
1. Match based on semantic similarity between table column headers and JSON dict keys.
2. Match based on context around the table (title, surrounding text).
3. Each table should be matched to at most ONE JSON key.
4. Each JSON key should be matched to at most ONE table.
5. If no good match exists for a table, omit it from the result.

Return a JSON array of objects, each with:
- "table_index": the table index number
- "json_key": the exact JSON key name (string) that matches this table
"""

    human_prompt = f"""DOCX Tables:
{json.dumps(tables_desc, indent=2, ensure_ascii=False)}

JSON Array Keys:
{json.dumps(keys_desc, indent=2, ensure_ascii=False)}

Return a JSON array matching each table_index to its json_key.
Output ONLY the JSON array, no markdown formatting.
"""

    response = llm.invoke(
        [
            SystemMessage(content=system_prompt),
            HumanMessage(content=human_prompt),
        ]
    )

    result = {}
    for item in _response_items(response.content):
        index, key = item.get("table_index"), item.get("json_key")
        if type(index) is int and index in valid_indices and isinstance(key, str) and key in array_keys:
            result[index] = key

    # Save to cache
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_data = {
        "inputs_hash": _compute_table_inputs_hash(data_tables, array_keys),
        "table_key_mapping": {str(k): v for k, v in result.items()},
    }
    existing_columns = _read_cache(TABLE_CACHE_FILE).get("column_mappings")
    if isinstance(existing_columns, dict):
        cache_data["column_mappings"] = existing_columns
    TABLE_CACHE_FILE.write_text(json.dumps(cache_data, indent=2, ensure_ascii=False), encoding="utf-8")

    for tidx, key in result.items():
        print(f"  Table {tidx} → '{key}'")

    return result


def map_table_columns_with_llm(
    table_index: int,
    headers: list[str],
    data_keys: list[str],
) -> dict[int, str]:
    """
    Use LLM to match table column headers to JSON dict keys.

    Returns:
        dict mapping column_index -> dict_key
    """
    if not headers or not data_keys:
        return {}

    # Build a cache key from headers + data_keys
    cache_id = hashlib.sha256(
        json.dumps(
            {"headers": headers, "data_keys": sorted(data_keys), "model": _model_name()}, sort_keys=True
        ).encode()
    ).hexdigest()[:16]

    cache = _read_cache(TABLE_CACHE_FILE)
    columns = cache.get("column_mappings", {})
    if isinstance(columns, dict) and isinstance(columns.get(cache_id), dict):
        try:
            cached = {int(key): value for key, value in columns[cache_id].items()}
            if all(0 <= key < len(headers) and value in data_keys for key, value in cached.items()):
                return cached
        except (TypeError, ValueError):
            pass

    headers_clean = [h.replace("\n", " ").strip() for h in headers]

    print(f"[LLM] Matching {len(headers_clean)} columns to {len(data_keys)} dict keys for table {table_index}...")

    llm = ChatOpenAI(model=_model_name(), temperature=0)

    system_prompt = """You are an expert at matching Romanian procurement document table columns to JSON data fields.

Given a list of DOCX table column headers and a list of JSON dict keys, match each column to the correct dict key.

RULES:
1. Match based on semantic similarity. Account for Romanian diacritics and morphology differences.
2. Each column should be matched to at most ONE dict key.
3. Each dict key should be matched to at most ONE column.
4. If no good match exists for a column, omit it.

Return a JSON array of objects, each with:
- "column_index": the 0-based column index
- "dict_key": the exact JSON dict key name (string)
"""

    human_prompt = f"""Table Column Headers (with their 0-based indices):
{json.dumps([{"index": i, "header": h} for i, h in enumerate(headers_clean)], indent=2, ensure_ascii=False)}

JSON Dict Keys:
{json.dumps(data_keys, indent=2, ensure_ascii=False)}

Return a JSON array matching each column_index to its dict_key.
Output ONLY the JSON array, no markdown formatting.
"""

    response = llm.invoke(
        [
            SystemMessage(content=system_prompt),
            HumanMessage(content=human_prompt),
        ]
    )

    result = {}
    for item in _response_items(response.content):
        index, key = item.get("column_index"), item.get("dict_key")
        if type(index) is int and 0 <= index < len(headers) and isinstance(key, str) and key in data_keys:
            result[index] = key

    # Save to cache
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = _read_cache(TABLE_CACHE_FILE)
    if not isinstance(cache.get("column_mappings"), dict):
        cache["column_mappings"] = {}
    cache["column_mappings"][cache_id] = {str(k): v for k, v in result.items()}
    TABLE_CACHE_FILE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    for cidx, key in sorted(result.items()):
        print(f"  Col {cidx} ({headers_clean[cidx][:30]}) → '{key}'")

    return result
