"""
Field Mapper — Uses LLM to determine the EXACT value to fill for each DOCX placeholder from JSON.

Sends the DOCX placeholder contexts + JSON data to GPT-4o-mini,
which returns the best string value for each placeholder. The mapping is
cached to .cache/mapping.json for deterministic subsequent runs.
"""

import hashlib
import json
from pathlib import Path

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI

from docx_analyzer import PlaceholderInfo, TableInfo

CACHE_DIR = Path(".cache")
CACHE_FILE = CACHE_DIR / "mapping.json"
TABLE_CACHE_FILE = CACHE_DIR / "table_mapping.json"


def _placeholder_id(p: PlaceholderInfo) -> str:
    """Generate a unique, stable ID for a placeholder based on its position."""
    # We include section/header_footer info to ensure uniqueness across the doc
    hf_prefix = f"s{p.section_index}_{p.header_footer_type}" if p.header_footer_type != "body" else "body"
    
    if p.location_type == "table":
        return f"{hf_prefix}_table_{p.table_index}_r{p.row_index}_c{p.cell_index}_p{p.cell_paragraph_index}"
    else:
        return f"{hf_prefix}_para_{p.paragraph_index}_run{p.run_index}_{p.placeholder_pattern[:10]}"


def _compute_inputs_hash(placeholders: list[PlaceholderInfo], json_data: dict) -> str:
    """Compute a hash of the inputs to detect changes that invalidate the cache."""
    # We now depend on the ACTUAL DATA in json_data, not just keys
    data = json.dumps({
        "placeholders": [_placeholder_id(p) for p in placeholders],
        "json_data": json_data,
        "version": "v2_with_used_keys",
    }, sort_keys=True)
    return hashlib.sha256(data.encode()).hexdigest()[:16]


def load_cached_mapping(placeholders: list[PlaceholderInfo], json_data: dict) -> dict[str, dict] | None:
    """Load mapping from cache if it exists and inputs haven't changed."""
    if not CACHE_FILE.exists():
        return None
    
    try:
        cache = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        # We now check against json_data content
        expected_hash = _compute_inputs_hash(placeholders, json_data)
        
        if cache.get("inputs_hash") == expected_hash:
            print("[Cache] Loaded mapping from .cache/mapping.json")
            return dict(cache["mapping"])
        else:
            print("[Cache] Inputs changed, regenerating mapping...")
            return None
    except (json.JSONDecodeError, KeyError):
        return None


def save_mapping_to_cache(mapping: dict, 
                          placeholders: list[PlaceholderInfo], 
                          json_data: dict) -> None:
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


def map_fields_with_llm(placeholders: list[PlaceholderInfo], 
                        json_keys: list[str],
                        json_data: dict) -> dict[str, dict]:
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
        model="gpt-5.1",
        #temperature=0,
        model_kwargs={"seed": 42},  # deterministic
    )
    
    # Split into batches to avoid token limits (5 placeholders per batch)
    BATCH_SIZE = 10
    all_mappings: dict[str, dict] = {}

    
    # We pass the full JSON object now, not just keys
    json_context = json.dumps(json_data, indent=2, ensure_ascii=False)
    
    for batch_start in range(0, len(descriptions), BATCH_SIZE):
        batch = descriptions[batch_start:batch_start + BATCH_SIZE]
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

        response = llm.invoke([
            SystemMessage(content=system_prompt),
            HumanMessage(content=human_prompt),
        ])
        
        # Parse response
        response_text = str(response.content).strip()
        # Remove markdown code fences if present
        if response_text.startswith("```"):
            response_text = response_text.split("\n", 1)[1]
            if response_text.endswith("```"):
                response_text = response_text[:-3].strip()
        
        try:
            batch_mappings = json.loads(response_text)
            for item in batch_mappings:
                all_mappings[item["id"]] = {
                    "value": item.get("value"),
                    "used_keys": item.get("used_json_keys", [])
                }
        except json.JSONDecodeError as e:
            print(f"  [LLM] Warning: Failed to parse batch {batch_num} response: {e}")
            print(f"  Response was: {response_text[:500]}")
            # Set all to None for this batch
            for desc in batch:
                all_mappings[desc["id"]] = {"value": None, "used_keys": []}
    
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
    data = json.dumps({
        "tables": [{"index": t.table_index, "headers": t.headers, "context": t.context} for t in data_tables],
        "array_keys": {k: str(type(v[0]).__name__) if v else "empty" for k, v in array_keys.items()},
        "version": "v1_llm_table_mapping",
    }, sort_keys=True)
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
    
    # Check cache
    if TABLE_CACHE_FILE.exists():
        try:
            cache = json.loads(TABLE_CACHE_FILE.read_text(encoding="utf-8"))
            expected_hash = _compute_table_inputs_hash(data_tables, array_keys)
            if cache.get("inputs_hash") == expected_hash and "table_key_mapping" in cache:
                print("[Cache] Loaded table-key mapping from cache")
                return {int(k): v for k, v in cache["table_key_mapping"].items()}
        except (json.JSONDecodeError, KeyError):
            pass
    
    print(f"[LLM] Matching {len(data_tables)} data tables to {len(array_keys)} JSON array keys...")
    
    # Build descriptions
    tables_desc = []
    for t in data_tables:
        headers_clean = [h.replace("\n", " ").strip() for h in t.headers if h.strip()]
        tables_desc.append({
            "table_index": t.table_index,
            "context": t.context[:300],
            "column_headers": headers_clean,
            "num_data_rows": len(t.data_row_indices),
        })
    
    keys_desc = {}
    for key, data in array_keys.items():
        if data and isinstance(data[0], dict):
            keys_desc[key] = {"type": "array_of_dicts", "dict_keys": list(data[0].keys()), "count": len(data)}
        elif data:
            keys_desc[key] = {"type": "array_of_strings", "sample": data[:3], "count": len(data)}
        else:
            keys_desc[key] = {"type": "empty_array", "count": 0}
    
    llm = ChatOpenAI(model="gpt-5.1", temperature=0)
    
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

    response = llm.invoke([
        SystemMessage(content=system_prompt),
        HumanMessage(content=human_prompt),
    ])
    
    response_text = str(response.content).strip()
    if response_text.startswith("```"):
        response_text = response_text.split("\n", 1)[1]
        if response_text.endswith("```"):
            response_text = response_text[:-3].strip()
    
    result = {}
    try:
        mappings = json.loads(response_text)
        for item in mappings:
            result[item["table_index"]] = item["json_key"]
    except (json.JSONDecodeError, KeyError) as e:
        print(f"[LLM] Warning: Failed to parse table-key mapping: {e}")
        print(f"  Response: {response_text[:500]}")
    
    # Save to cache
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_data = {
        "inputs_hash": _compute_table_inputs_hash(data_tables, array_keys),
        "table_key_mapping": {str(k): v for k, v in result.items()},
    }
    # Merge with existing column mappings if present
    if TABLE_CACHE_FILE.exists():
        try:
            existing = json.loads(TABLE_CACHE_FILE.read_text(encoding="utf-8"))
            if "column_mappings" in existing:
                cache_data["column_mappings"] = existing["column_mappings"]
        except (json.JSONDecodeError, KeyError):
            pass
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
        json.dumps({"headers": headers, "data_keys": sorted(data_keys)}, sort_keys=True).encode()
    ).hexdigest()[:16]
    
    # Check cache
    if TABLE_CACHE_FILE.exists():
        try:
            cache = json.loads(TABLE_CACHE_FILE.read_text(encoding="utf-8"))
            col_cache = cache.get("column_mappings", {})
            if cache_id in col_cache:
                print(f"[Cache] Loaded column mapping for table {table_index} from cache")
                return {int(k): v for k, v in col_cache[cache_id].items()}
        except (json.JSONDecodeError, KeyError):
            pass
    
    headers_clean = [h.replace("\n", " ").strip() for h in headers]
    
    print(f"[LLM] Matching {len(headers_clean)} columns to {len(data_keys)} dict keys for table {table_index}...")
    
    llm = ChatOpenAI(model="gpt-5.1", temperature=0)
    
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

    response = llm.invoke([
        SystemMessage(content=system_prompt),
        HumanMessage(content=human_prompt),
    ])
    
    response_text = str(response.content).strip()
    if response_text.startswith("```"):
        response_text = response_text.split("\n", 1)[1]
        if response_text.endswith("```"):
            response_text = response_text[:-3].strip()
    
    result = {}
    try:
        mappings = json.loads(response_text)
        for item in mappings:
            result[item["column_index"]] = item["dict_key"]
    except (json.JSONDecodeError, KeyError) as e:
        print(f"[LLM] Warning: Failed to parse column mapping: {e}")
        print(f"  Response: {response_text[:500]}")
    
    # Save to cache
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if TABLE_CACHE_FILE.exists():
        try:
            cache = json.loads(TABLE_CACHE_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, KeyError):
            cache = {}
    else:
        cache = {}
    
    if "column_mappings" not in cache:
        cache["column_mappings"] = {}
    cache["column_mappings"][cache_id] = {str(k): v for k, v in result.items()}
    TABLE_CACHE_FILE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    
    for cidx, key in sorted(result.items()):
        print(f"  Col {cidx} ({headers_clean[cidx][:30]}) → '{key}'")
    
    return result
