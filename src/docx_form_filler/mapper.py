"""Semantic mapping of document blanks and tables to source JSON, via the chat model.

The model proposes decisions; this module validates their shape against the
requested IDs and the source keys and caches accepted decisions so unchanged
inputs give repeatable results. Validation cannot prove a value is correct.
"""

from __future__ import annotations

import logging
from collections.abc import Mapping, Sequence
from dataclasses import asdict
from pathlib import Path
from typing import Any, TypedDict, cast

from docx_form_filler import llm
from docx_form_filler.analyzer import PlaceholderInfo, TableInfo, placeholder_id
from docx_form_filler.cache import JsonCache, fingerprint

logger = logging.getLogger(__name__)

CACHE_DIR = Path(".cache")
FIELD_CACHE = JsonCache(CACHE_DIR / "mapping.json")
TABLE_CACHE = JsonCache(CACHE_DIR / "table_mapping.json")

FIELD_BATCH_SIZE = 10
CHECKBOX_VALUES = frozenset({"CHECKED", "UNCHECKED"})
# Bump when prompts or the cached decision format change.
FIELD_CACHE_VERSION = "v4"
TABLE_CACHE_VERSION = "v3"


class FieldDecision(TypedDict):
    value: str | None
    used_keys: list[str]


FieldMapping = dict[str, FieldDecision]


def map_fields(placeholders: Sequence[PlaceholderInfo], json_data: Mapping[str, Any]) -> FieldMapping:
    """Decide the value for every placeholder; undecided fields map to ``None``."""
    cached = load_cached_mapping(placeholders, json_data)
    if cached is not None:
        logger.info("Using cached field mapping for %d placeholders", len(cached))
        return cached

    checkboxes = {placeholder_id(field) for field in placeholders if field.is_checkbox}
    descriptions = [_describe_placeholder(field) for field in placeholders]
    source = llm.to_json(json_data)
    mapping: FieldMapping = {}
    batches = range(0, len(descriptions), FIELD_BATCH_SIZE)
    for number, start in enumerate(batches, start=1):
        batch = descriptions[start : start + FIELD_BATCH_SIZE]
        logger.info("Mapping placeholder batch %d/%d (%d fields)", number, len(batches), len(batch))
        requested = {item["id"] for item in batch}
        mapping.update({pid: {"value": None, "used_keys": []} for pid in requested})
        replies = llm.request_json_objects(
            "fields",
            f"JSON data source:\n{source}\n\nDOCX placeholders to fill:\n{llm.to_json(batch)}",
        )
        for item in replies:
            pid, value, keys = item.get("id"), item.get("value"), item.get("used_json_keys")
            is_checkbox = pid in checkboxes
            if isinstance(pid, str) and pid in requested and _is_valid_decision(value, keys, is_checkbox, json_data):
                mapping[pid] = {"value": value, "used_keys": cast(list[str], keys)}

    save_mapping_to_cache(mapping, placeholders, json_data)
    filled = sum(decision["value"] is not None for decision in mapping.values())
    logger.info("Model proposed values for %d of %d placeholders", filled, len(mapping))
    return mapping


def map_tables_to_arrays(tables: Sequence[TableInfo], arrays: Mapping[str, list[Any]]) -> dict[int, str]:
    """Match each data table to at most one JSON array key."""
    if not tables or not arrays:
        return {}
    valid_tables = {table.table_index for table in tables}
    key = _table_inputs_hash(tables, arrays)
    cache = TABLE_CACHE.read()
    if cache.get("inputs_hash") == key:
        cached = _int_keyed(cache.get("table_key_mapping"))
        if cached is not None and all(i in valid_tables and v in arrays for i, v in cached.items()):
            return cached

    logger.info("Matching %d data tables to %d JSON arrays", len(tables), len(arrays))
    replies = llm.request_json_objects(
        "tables",
        f"DOCX tables:\n{llm.to_json([_describe_table(t) for t in tables])}\n\n"
        f"JSON array keys:\n{llm.to_json({name: _describe_array(rows) for name, rows in arrays.items()})}",
    )
    result: dict[int, str] = {}
    for item in replies:
        index, name = item.get("table_index"), item.get("json_key")
        if type(index) is int and index in valid_tables and isinstance(name, str) and name in arrays:
            result[index] = name

    cache = TABLE_CACHE.read()
    cache.update(inputs_hash=key, table_key_mapping={str(i): v for i, v in result.items()})
    TABLE_CACHE.write(cache)
    for index, name in result.items():
        logger.info("Table %d -> %r", index, name)
    return result


def map_table_columns(table_index: int, headers: Sequence[str], data_keys: Sequence[str]) -> dict[int, str]:
    """Match the columns of one table to the keys of its array's objects."""
    if not headers or not data_keys:
        return {}
    key = fingerprint({"headers": list(headers), "data_keys": sorted(data_keys), "model": llm.model_name()})
    cache = TABLE_CACHE.read()
    columns = cache.get("column_mappings")
    if isinstance(columns, dict):
        cached = _int_keyed(columns.get(key))
        if cached is not None and all(0 <= i < len(headers) and v in data_keys for i, v in cached.items()):
            return cached

    clean_headers = [header.replace("\n", " ").strip() for header in headers]
    logger.info("Matching %d columns of table %d to %d keys", len(clean_headers), table_index, len(data_keys))
    replies = llm.request_json_objects(
        "columns",
        "Table column headers (with their 0-based indices):\n"
        f"{llm.to_json([{'index': i, 'header': h} for i, h in enumerate(clean_headers)])}\n\n"
        f"JSON object keys:\n{llm.to_json(list(data_keys))}",
    )
    result: dict[int, str] = {}
    for item in replies:
        index, name = item.get("column_index"), item.get("dict_key")
        if type(index) is int and 0 <= index < len(headers) and isinstance(name, str) and name in data_keys:
            result[index] = name

    cache = TABLE_CACHE.read()
    if not isinstance(cache.get("column_mappings"), dict):
        cache["column_mappings"] = {}
    cache["column_mappings"][key] = {str(i): v for i, v in result.items()}
    TABLE_CACHE.write(cache)
    for index, name in sorted(result.items()):
        logger.info("Column %d (%s) -> %r", index, clean_headers[index][:30], name)
    return result


def load_cached_mapping(placeholders: Sequence[PlaceholderInfo], json_data: Mapping[str, Any]) -> FieldMapping | None:
    """Return the cached field mapping if it matches these inputs and is well formed."""
    cache = FIELD_CACHE.read()
    mapping = cache.get("mapping")
    if cache.get("inputs_hash") != _field_inputs_hash(placeholders, json_data) or not isinstance(mapping, dict):
        return None
    fields = {placeholder_id(field): field for field in placeholders}
    for pid, decision in mapping.items():
        if pid not in fields or not isinstance(decision, dict):
            return None
        if not _is_valid_decision(decision.get("value"), decision.get("used_keys"), fields[pid].is_checkbox, json_data):
            return None
    return mapping


def save_mapping_to_cache(
    mapping: FieldMapping, placeholders: Sequence[PlaceholderInfo], json_data: Mapping[str, Any]
) -> None:
    FIELD_CACHE.write({"inputs_hash": _field_inputs_hash(placeholders, json_data), "mapping": mapping})
    logger.debug("Saved field mapping to %s", FIELD_CACHE.path)


def _is_valid_decision(value: object, keys: object, is_checkbox: bool, json_data: Mapping[str, Any]) -> bool:
    """Shape check shared by model replies and cache entries."""
    if value is not None and (not isinstance(value, str) or (is_checkbox and value not in CHECKBOX_VALUES)):
        return False
    return isinstance(keys, list) and all(isinstance(key, str) and key in json_data for key in keys)


def _int_keyed(value: object) -> dict[int, str] | None:
    """Decode a JSON object with integer-string keys and string values."""
    if not isinstance(value, dict):
        return None
    try:
        result = {int(key): item for key, item in value.items()}
    except (TypeError, ValueError):
        return None
    return result if all(isinstance(item, str) for item in result.values()) else None


def _field_inputs_hash(placeholders: Sequence[PlaceholderInfo], json_data: Mapping[str, Any]) -> str:
    return fingerprint(
        {
            "placeholders": [asdict(field) for field in placeholders],
            "json_data": json_data,
            "version": FIELD_CACHE_VERSION,
            "model": llm.model_name(),
        }
    )


def _table_inputs_hash(tables: Sequence[TableInfo], arrays: Mapping[str, list[Any]]) -> str:
    return fingerprint(
        {
            "tables": [asdict(table) for table in tables],
            "array_keys": arrays,
            "version": TABLE_CACHE_VERSION,
            "model": llm.model_name(),
        }
    )


def _describe_placeholder(field: PlaceholderInfo) -> dict[str, Any]:
    pid = placeholder_id(field)
    if field.is_checkbox:
        return {
            "id": pid,
            "type": "checkbox",
            "label": field.checkbox_label,
            "full_context": field.full_paragraph_text[:300],
            "nearby_paragraphs": field.surrounding_context[:300],
        }
    if field.location_type == "table":
        return {
            "id": pid,
            "type": "table_cell",
            "column_header": field.column_header,
            "context_before": field.context_before[:150],
            "context_after": field.context_after[:150],
            "full_context": field.full_paragraph_text[:300],
        }
    return {
        "id": pid,
        "type": "text_blank",
        "context_before": field.context_before[:200],
        "context_after": field.context_after[:200],
        "full_context": field.full_paragraph_text[:800],
        "nearby_paragraphs": field.surrounding_context[:800],
    }


def _describe_table(table: TableInfo) -> dict[str, Any]:
    return {
        "table_index": table.table_index,
        "context": table.context[:300],
        "column_headers": [header.replace("\n", " ").strip() for header in table.headers if header.strip()],
        "num_data_rows": len(table.data_row_indices),
    }


def _describe_array(rows: list[Any]) -> dict[str, Any]:
    if rows and isinstance(rows[0], dict):
        return {"type": "array_of_dicts", "dict_keys": object_keys(rows), "count": len(rows)}
    if rows:
        return {"type": "array_of_strings", "sample": rows[:3], "count": len(rows)}
    return {"type": "empty_array", "count": 0}


def object_keys(rows: Sequence[object]) -> list[str]:
    """Keys of all object rows, in first-seen order (later rows may add fields)."""
    return list(dict.fromkeys(key for row in rows if isinstance(row, dict) for key in row))
