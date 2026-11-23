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
import sys
from pathlib import Path

from docx import Document
from dotenv import load_dotenv

import docx_filler
from docx_analyzer import get_all_placeholders, save_analysis_report
from field_mapper import (
    _placeholder_id,
    map_fields_with_llm,
    map_table_columns_with_llm,
    map_table_to_json_key_with_llm,
)


def main():
    load_dotenv()

    parser = argparse.ArgumentParser(description="Fill a DOCX form using JSON data and semantic field mapping.")
    parser.add_argument("docx_path", type=Path, help="Path to the source DOCX template")
    parser.add_argument("json_path", type=Path, help="Path to a JSON object with source data")
    args = parser.parse_args()
    docx_path = args.docx_path
    json_path = args.json_path

    if not docx_path.is_file():
        print(f"Error: DOCX file not found: {docx_path}")
        sys.exit(1)
    if not json_path.is_file():
        print(f"Error: JSON file not found: {json_path}")
        sys.exit(1)

    # --- Load inputs ---
    print(f"[Main] Loading DOCX: {docx_path}")
    doc = Document(str(docx_path))

    print(f"[Main] Loading JSON: {json_path}")
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)
    except (OSError, ValueError) as error:
        print(f"Error: Invalid JSON input: {error}")
        sys.exit(1)

    if not isinstance(json_data, dict):
        print("Error: JSON root must be an object")
        sys.exit(1)

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
                            list(arr[0].keys()),
                        )
                        column_mappings[table_info.table_index] = col_map

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

    # --- Statistics & Logging ---
    all_used_keys = mapper_used_keys.union(table_used_keys)
    unused_keys = [k for k in json_keys if k not in all_used_keys]

    # Determine unfilled placeholders
    unfilled_placeholders = []
    for p in placeholders:
        pid = _placeholder_id(p)
        val_obj = mapping.get(pid)
        if not val_obj or val_obj.get("value") is None:
            unfilled_placeholders.append(p)

    # Determine unfilled data tables
    filled_table_indices = {t_stat["table_index"] for t_stat in table_stats}
    unfilled_tables = [t for t in data_tables if t.table_index not in filled_table_indices]

    print("\n=== FINAL STATISTICS ===")
    print(f"✅ Placeholders Filled: {filled_count}/{len(placeholders)}")

    total_rows_filled = 0
    for t_stat in table_stats:
        print(f"📋 Table {t_stat['table_index']} ({t_stat['json_key']}): {t_stat['rows_filled']} rows filled")
        total_rows_filled += t_stat["rows_filled"]

    if unfilled_tables:
        print(f"\n❌ Unfilled Data Tables: {len(unfilled_tables)}")
        for t in unfilled_tables:
            print(f"   - Table {t.table_index}: headers={t.headers[:3]}{'...' if len(t.headers) > 3 else ''}")

    if unfilled_placeholders:
        print(f"\n❌ Unfilled Placeholders: {len(unfilled_placeholders)}/{len(placeholders)}")
        for p in unfilled_placeholders[:10]:
            ctx = (p.context_before[-30:] + p.placeholder_pattern + p.context_after[:30]).strip()
            print(f"   - [{p.location_type}] '{p.placeholder_pattern}' | {ctx[:80]}")
        if len(unfilled_placeholders) > 10:
            print(f"   ... and {len(unfilled_placeholders) - 10} more (see docx_full_log.txt for full list)")

    print(f"\n🔑 JSON Fields Used: {len(all_used_keys)}/{len(json_keys)}")
    print(f"⚠️ Unused JSON Fields: {len(unused_keys)}")
    if unused_keys:
        print(f"   (First 5 unused: {', '.join(unused_keys[:5])}...)")

    # Append stats and DETAILED LOG to log file
    with open("docx_full_log.txt", "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 60 + "\n")
        f.write("FILLING STATISTICS\n")
        f.write("=" * 60 + "\n")
        f.write(f"Placeholders Filled: {filled_count}/{len(placeholders)}\n")
        f.write(f"Total Table Rows Filled: {total_rows_filled}\n")
        for t_stat in table_stats:
            f.write(f"  - Table {t_stat['table_index']} ({t_stat['json_key']}): {t_stat['rows_filled']} rows\n")

        # Unfilled data tables
        f.write("\n" + "-" * 40 + "\n")
        f.write(f"UNFILLED DATA TABLES: {len(unfilled_tables)}\n")
        if unfilled_tables:
            for t in unfilled_tables:
                f.write(f"  - Table {t.table_index}\n")
                f.write(f"    Headers: {t.headers}\n")
                f.write(f"    Data row indices: {t.data_row_indices}\n")
                f.write(f"    Context: {t.context[:200]}\n")
        else:
            f.write("  (all data tables were filled)\n")

        # Unfilled placeholders — full list
        f.write("\n" + "-" * 40 + "\n")
        f.write(f"UNFILLED PLACEHOLDERS: {len(unfilled_placeholders)}/{len(placeholders)}\n")
        if unfilled_placeholders:
            for i, p in enumerate(unfilled_placeholders):
                pid = _placeholder_id(p)
                ctx = (p.context_before[-40:] + " [" + p.placeholder_pattern + "] " + p.context_after[:40]).strip()
                f.write(f"  {i + 1}. ID: {pid}\n")
                f.write(f"     Location: {p.location_type} | Section: {p.section_index} | HF: {p.header_footer_type}\n")
                f.write(f"     Pattern: '{p.placeholder_pattern}'\n")
                f.write(f"     Context: {ctx[:120]}\n")
                if p.location_type == "table":
                    f.write(f"     Table: {p.table_index}, Row: {p.row_index}, Cell: {p.cell_index}\n")
                    f.write(f"     Column Header: '{p.column_header}'\n")
                elif p.location_type == "paragraph":
                    f.write(f"     Paragraph: {p.paragraph_index}, Run: {p.run_index}\n")
                f.write(f"     Surrounding: {p.surrounding_context[:120]}\n")
        else:
            f.write("  (all placeholders were filled)\n")

        f.write("\n" + "-" * 40 + "\n")
        f.write("JSON USAGE\n")
        f.write(f"Total Fields: {len(json_keys)}\n")
        f.write(f"Used Fields:  {len(all_used_keys)}\n")
        f.write(f"Unused Fields: {len(unused_keys)}\n")
        if unused_keys:
            f.write("Unused Keys List:\n")
            for k in unused_keys:
                f.write(f"  - {k}\n")

        # DETAILED LOGGING
        f.write("\n\n" + "=" * 60 + "\n")
        f.write("DETAILED FILL LOG (Item by Item)\n")
        f.write("=" * 60 + "\n")

        # Log Placeholders
        f.write("\n[PLACEHOLDERS]\n")
        for log in fill_log:
            if log["type"] == "placeholder" or log["type"] == "placeholder_table":
                # Format: [P] {id} | Val: "{val}" | Src: {src}
                line = f"[P] ID: {log['id']} | Value: {log['value']!r} | Source: {log['source']}"
                if "context" in log:
                    line += f" | Context: {log['context'].strip()[:50]}..."
                f.write(line + "\n")

        # Log Table Cells
        f.write("\n[TABLE DATA]\n")
        for log in fill_log:
            if log["type"] == "table_cell":
                # Format: [T] T{idx} R{r} C{c} | Val: "{val}" | Src: {src}
                line = f"[T] Table {log['table_idx']} Row {log['row']} Col {log['col']} | Value: {log['value']!r} | Source: {log['source']}"
                f.write(line + "\n")

    # --- Step 4: Save output ---
    output_dir = Path("resources")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_name = f"{docx_path.stem}.filled.docx"
    output_path = output_dir / output_name

    doc.save(str(output_path))
    print(f"\n[Main] ✅ Saved filled document: {output_path}")
    print(f"[Main] Original: {docx_path}")
    print(f"[Main] Output:   {output_path}")


if __name__ == "__main__":
    main()
