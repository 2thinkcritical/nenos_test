"""Fill diagnostics: a short console summary and a detailed plain-text report."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from docx_form_filler.analyzer import PlaceholderInfo, TableInfo, placeholder_id
from docx_form_filler.filler import FillResult
from docx_form_filler.mapper import FieldMapping

SUMMARY_LIST_LIMIT = 10


@dataclass
class FillSummary:
    placeholders: Sequence[PlaceholderInfo]
    unfilled_placeholders: list[PlaceholderInfo]
    unfilled_tables: list[TableInfo]
    json_keys: list[str]
    used_keys: set[str]
    unused_keys: list[str]
    fill: FillResult

    @property
    def table_rows_filled(self) -> int:
        return sum(stat["rows_filled"] for stat in self.fill.table_stats)


def summarize(
    placeholders: Sequence[PlaceholderInfo],
    data_tables: Sequence[TableInfo],
    mapping: FieldMapping,
    json_keys: Sequence[str],
    fill: FillResult,
) -> FillSummary:
    decisions = [mapping.get(placeholder_id(field)) for field in placeholders]
    used = {key for decision in decisions if decision for key in decision.get("used_keys", [])}
    used |= fill.table_used_keys
    filled_tables = {stat["table_index"] for stat in fill.table_stats}
    return FillSummary(
        placeholders=placeholders,
        unfilled_placeholders=[
            field
            for field, decision in zip(placeholders, decisions, strict=True)
            if decision is None or decision.get("value") is None
        ],
        unfilled_tables=[table for table in data_tables if table.table_index not in filled_tables],
        json_keys=list(json_keys),
        used_keys=used,
        unused_keys=[key for key in json_keys if key not in used],
        fill=fill,
    )


def format_summary(summary: FillSummary, report_path: Path) -> str:
    """Human-readable result for the terminal."""
    total = len(summary.placeholders)
    lines = [f"Placeholders filled: {summary.fill.filled_count}/{total}"]
    lines += [
        f"Table {stat['table_index']} ({stat['json_key']}): {stat['rows_filled']} rows filled"
        for stat in summary.fill.table_stats
    ]
    if summary.unfilled_tables:
        lines.append(f"Unfilled data tables: {len(summary.unfilled_tables)}")
        lines += [f"  - table {table.table_index}: headers={table.headers[:3]}" for table in summary.unfilled_tables]
    if summary.unfilled_placeholders:
        lines.append(f"Unfilled placeholders: {len(summary.unfilled_placeholders)}/{total}")
        for field in summary.unfilled_placeholders[:SUMMARY_LIST_LIMIT]:
            context = (field.context_before[-30:] + field.placeholder_pattern + field.context_after[:30]).strip()
            lines.append(f"  - [{field.location_type}] {context[:80]}")
        hidden = len(summary.unfilled_placeholders) - SUMMARY_LIST_LIMIT
        if hidden > 0:
            lines.append(f"  ... and {hidden} more (see {report_path})")
    lines.append(f"JSON fields used: {len(summary.used_keys)}/{len(summary.json_keys)}")
    if summary.unused_keys:
        lines.append(f"Unused JSON fields: {', '.join(summary.unused_keys)}")
    return "\n".join(lines)


def append_report(summary: FillSummary, path: Path) -> None:
    """Append statistics and the value-level fill log to the analysis report."""
    rule, thin = "=" * 60, "-" * 40
    total = len(summary.placeholders)
    lines = ["", "", rule, "FILLING STATISTICS", rule]
    lines.append(f"Placeholders filled: {summary.fill.filled_count}/{total}")
    lines.append(f"Total table rows filled: {summary.table_rows_filled}")
    lines += [
        f"  - Table {stat['table_index']} ({stat['json_key']}): {stat['rows_filled']} rows"
        for stat in summary.fill.table_stats
    ]

    lines += ["", thin, f"UNFILLED DATA TABLES: {len(summary.unfilled_tables)}"]
    for table in summary.unfilled_tables:
        lines += [
            f"  - Table {table.table_index}",
            f"    Headers: {table.headers}",
            f"    Data row indices: {table.data_row_indices}",
            f"    Context: {table.context[:200]}",
        ]
    if not summary.unfilled_tables:
        lines.append("  (all data tables were filled)")

    lines += ["", thin, f"UNFILLED PLACEHOLDERS: {len(summary.unfilled_placeholders)}/{total}"]
    for number, field in enumerate(summary.unfilled_placeholders, start=1):
        context = f"{field.context_before[-40:]} [{field.placeholder_pattern}] {field.context_after[:40]}".strip()
        lines += [
            f"  {number}. ID: {placeholder_id(field)}",
            f"     Location: {field.location_type} | Section: {field.section_index} | HF: {field.header_footer_type}",
            f"     Pattern: {field.placeholder_pattern!r}",
            f"     Context: {context[:120]}",
        ]
        if field.location_type == "table":
            lines += [
                f"     Table: {field.table_index}, Row: {field.row_index}, Cell: {field.cell_index}",
                f"     Column header: {field.column_header!r}",
            ]
        else:
            lines.append(f"     Paragraph: {field.paragraph_index}, Run: {field.run_index}")
        lines.append(f"     Surrounding: {field.surrounding_context[:120]}")
    if not summary.unfilled_placeholders:
        lines.append("  (all placeholders were filled)")

    lines += [
        "",
        thin,
        "JSON USAGE",
        f"Total fields: {len(summary.json_keys)}",
        f"Used fields:  {len(summary.used_keys)}",
        f"Unused fields: {len(summary.unused_keys)}",
    ]
    lines += [f"  - {key}" for key in summary.unused_keys]

    lines += ["", "", rule, "DETAILED FILL LOG", rule, "", "[PLACEHOLDERS]"]
    for entry in summary.fill.log:
        if entry["type"] in {"placeholder", "placeholder_table"}:
            line = f"[P] ID: {entry['id']} | Value: {entry['value']!r} | Source: {entry['source']}"
            if "context" in entry:
                line += f" | Context: {entry['context'].strip()[:50]}..."
            lines.append(line)
    lines += ["", "[TABLE DATA]"]
    lines += [
        f"[T] Table {entry['table_idx']} Row {entry['row']} Col {entry['col']} | "
        f"Value: {entry['value']!r} | Source: {entry['source']}"
        for entry in summary.fill.log
        if entry["type"] == "table_cell"
    ]
    with path.open("a", encoding="utf-8") as report:
        report.write("\n".join(lines) + "\n")
