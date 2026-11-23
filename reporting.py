"""Human-readable document fill diagnostics, separate from orchestration."""

from pathlib import Path

from docx_analyzer import PlaceholderInfo, TableInfo
from docx_analyzer import placeholder_id as _placeholder_id


def write_fill_report(
    *,
    placeholders: list[PlaceholderInfo],
    data_tables: list[TableInfo],
    mapping: dict[str, dict],
    json_keys: list[str],
    mapper_used_keys: set[str],
    table_used_keys: set[str],
    filled_count: int,
    table_stats: list[dict],
    fill_log: list[dict],
    report_path: Path,
) -> None:
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
    with report_path.open("a", encoding="utf-8") as f:
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
