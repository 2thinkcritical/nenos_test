"""
DOCX Form Filler — Agentic AI

Main entry point that orchestrates the form filling pipeline:
  1. Analyze the DOCX to find all placeholders
  2. Use LLM to map JSON keys to placeholder positions
  3. Fill the document in-place
  4. Save as <original_name>.filled.docx

Usage:
    uv run python main.py <docx_path> <json_path>
    uv run python main.py "sample_forms (1).docx" input_date.json
"""

import argparse
import json
import os
import sys
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from dotenv import load_dotenv
from lxml.etree import XMLSyntaxError
from openai import OpenAIError

import docx_filler
from docx_analyzer import get_all_placeholders, save_analysis_report
from field_mapper import (
    _placeholder_id,
    map_fields_with_llm,
    map_table_columns_with_llm,
    map_table_to_json_key_with_llm,
)
from reporting import write_fill_report


def run_pipeline(docx_path: Path, json_path: Path, output_path: Path | None = None) -> Path:
    """Fill a copy of a template; callers choose the working directory for reports/cache."""
    output_path = output_path or Path("resources") / f"{docx_path.stem}.filled.docx"
    resolved_output = output_path.resolve()
    if resolved_output == docx_path.resolve() or (
        output_path.exists() and docx_path.exists() and output_path.samefile(docx_path)
    ):
        raise ValueError("Output must not overwrite the input DOCX template")
    if output_path.suffix.lower() != ".docx":
        raise ValueError("Output path must end with .docx")
    if not docx_path.is_file():
        raise ValueError(f"DOCX file not found: {docx_path}")
    if not json_path.is_file():
        raise ValueError(f"JSON file not found: {json_path}")

    # --- Load inputs ---
    print(f"[Main] Loading DOCX: {docx_path}")
    try:
        doc = Document(str(docx_path))
    except (PackageNotFoundError, BadZipFile, KeyError, XMLSyntaxError, ValueError) as error:
        raise ValueError(f"Cannot read DOCX: {docx_path}. Provide a valid Word .docx file.") from error

    print(f"[Main] Loading JSON: {json_path}")
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)
    except (OSError, ValueError) as error:
        raise ValueError(f"Invalid JSON input: {error}") from error

    if not isinstance(json_data, dict):
        raise ValueError("JSON root must be an object")

    json_keys = list(json_data.keys())
    print(f"[Main] JSON contains {len(json_keys)} fields")

    # --- Step 1: Analyze the DOCX ---
    print("\n=== Step 1: Analyzing DOCX placeholders ===")
    _, placeholders, data_tables = get_all_placeholders(doc)

    # Save detailed report
    save_analysis_report(placeholders, data_tables, "docx_full_log.txt")

    print(f"Found {len(placeholders)} placeholders.")
    if data_tables:
        print(f"Found {len(data_tables)} data tables.")
    print(f"[Main] Found {len(placeholders)} placeholders and {len(data_tables)} data tables")

    # --- Step 2: Map fields using LLM ---
    print("\n=== Step 2: Mapping fields with LLM ===")

    # map_fields_with_llm now returns dict[id, {value, used_keys}]
    mapping = map_fields_with_llm(placeholders, json_keys, json_data)

    # Extract used keys from mapping
    mapper_used_keys = set()
    for v in mapping.values():
        mapper_used_keys.update(v.get("used_keys", []))

    print("\n--- Mapping Preview ---")
    print("\n[Main] Mapping summary:")
    for p in placeholders:
        pid = _placeholder_id(p)
        val_obj = mapping.get(pid)
        if val_obj and val_obj.get("value"):
            ctx = p.context_before[-40:] if p.context_before else ""
            print(
                f"  {ctx}{'...' if len(p.context_before) > 40 else ''} → {val_obj['value']} (Source: {val_obj['used_keys']})"
            )

    # --- Step 2.5: Map data tables using LLM ---
    table_key_mapping = {}
    column_mappings = {}

    if data_tables:
        print("\n=== Step 2.5: Mapping data tables with LLM ===")

        # Find JSON keys that contain array data
        array_keys = {}
        for key, value in json_data.items():
            if isinstance(value, list):
                array_keys[key] = value
            elif isinstance(value, str) and value.strip().startswith("["):
                try:
                    import json as _json

                    parsed = _json.loads(value)
                    if isinstance(parsed, list):
                        array_keys[key] = parsed
                except json.JSONDecodeError:
                    continue

        if array_keys:
            # Step 1: Match tables to JSON array keys
            table_key_mapping = map_table_to_json_key_with_llm(data_tables, array_keys)

            # Step 2: For each matched table with dict data, map columns
            for table_info in data_tables:
                matched_key = table_key_mapping.get(table_info.table_index)
                if matched_key and matched_key in array_keys:
                    arr = array_keys[matched_key]
                    if arr and isinstance(arr[0], dict):
                        col_map = map_table_columns_with_llm(
                            table_info.table_index,
                            table_info.headers,
                            list(dict.fromkeys(key for row in arr if isinstance(row, dict) for key in row)),
                        )
                        column_mappings[table_info.table_index] = col_map

    placeholders = docx_filler.exclude_array_placeholders(placeholders, data_tables, json_data, table_key_mapping)
    mapper_used_keys = {
        key for field in placeholders for key in mapping.get(_placeholder_id(field), {}).get("used_keys", [])
    }

    # --- Step 3: Fill the Document ---
    print("\n=== Step 3: Filling Document ===")
    # fill_document now returns fill_log
    filled_count, table_stats, table_used_keys, fill_log = docx_filler.fill_document(
        doc,
        placeholders,
        mapping,
        json_data,
        data_tables,
        table_key_mapping=table_key_mapping,
        column_mappings=column_mappings,
    )

    write_fill_report(
        placeholders=placeholders,
        data_tables=data_tables,
        mapping=mapping,
        json_keys=json_keys,
        mapper_used_keys=mapper_used_keys,
        table_used_keys=table_used_keys,
        filled_count=filled_count,
        table_stats=table_stats,
        fill_log=fill_log,
        report_path=Path("docx_full_log.txt"),
    )

    # --- Step 4: Save output ---
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=resolved_output.parent, prefix=".docx-fill-", suffix=".tmp", delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        doc.save(str(temporary_path))
        os.replace(temporary_path, resolved_output)
    finally:
        temporary_path.unlink(missing_ok=True)
    print(f"\n[Main] ✅ Saved filled document: {output_path}")
    print(f"[Main] Original: {docx_path}")
    print(f"[Main] Output:   {output_path}")
    return output_path


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Fill a DOCX form using JSON data and semantic field mapping.")
    parser.add_argument("docx_path", type=Path, help="Path to the source DOCX template")
    parser.add_argument("json_path", type=Path, help="Path to a JSON object with source data")
    parser.add_argument("--output", type=Path, help="Output DOCX path; defaults to resources/<template>.filled.docx")
    args = parser.parse_args()
    try:
        run_pipeline(args.docx_path, args.json_path, args.output)
    except OpenAIError:
        print(
            "Error: Model request failed. Check your API key, model access, network, and usage limits. No completed document was saved."
        )
        sys.exit(1)
    except (ValueError, OSError) as error:
        print(f"Error: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
