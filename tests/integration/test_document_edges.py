import pytest
from docx import Document
from docx.oxml import OxmlElement

from docx_analyzer import TableInfo, get_all_placeholders
from docx_filler import fill_document
from field_mapper import _placeholder_id


def test_hyperlink_text_does_not_shift_blank_replacement(tmp_path):
    doc = Document()
    paragraph = doc.add_paragraph()
    hyperlink = OxmlElement("w:hyperlink")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Instructions"
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    paragraph.add_run(" Name: ______")
    _, fields, _ = get_all_placeholders(doc)
    mapping = {_placeholder_id(fields[0]): {"value": "DEMO Alex", "used_keys": ["name"]}}
    fill_document(doc, fields, mapping, {"name": "DEMO Alex"})
    output = tmp_path / "filled.docx"
    doc.save(output)
    assert Document(output).paragraphs[0].text == "Instructions Name: DEMO Alex"


@pytest.mark.parametrize(
    "rows,expected",
    [
        (["First", "Second", "Third", "Fourth"], [("First", ""), ("Second", ""), ("Third", ""), ("Fourth", "")]),
        (
            [{"left": "First", "right": "A"}, {"left": "Second", "right": "B"}, {"left": "Third"}],
            [("First", "A"), ("Second", "B"), ("Third", "")],
        ),
    ],
)
def test_overflow_rows_do_not_inherit_previous_values(tmp_path, rows, expected):
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "left"
    table.cell(0, 1).text = "right"
    info = TableInfo(0, ["left", "right"], [1, 2], "Items")
    fill_document(doc, [], {}, {"items": rows}, [info], table_key_mapping={0: "items"})
    output = tmp_path / "rows.docx"
    doc.save(output)
    reopened = Document(output).tables[0]
    assert [tuple(cell.text for cell in row.cells) for row in reopened.rows[1:]] == expected


def test_header_blanks_are_scanned_once():
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "Name: ______"
    _, fields, _ = get_all_placeholders(doc)
    assert len(fields) == 1


def test_overflow_rows_are_inserted_before_totals(tmp_path):
    doc = Document()
    table = doc.add_table(rows=4, cols=2)
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Amount"
    table.cell(3, 0).text = "TOTAL"
    table.cell(3, 1).text = "60"
    info = TableInfo(0, ["Item", "Amount"], [1, 2], "Items")
    rows = [{"Item": "A", "Amount": 10}, {"Item": "B", "Amount": 20}, {"Item": "C", "Amount": 30}]
    fill_document(doc, [], {}, {"items": rows}, [info], table_key_mapping={0: "items"})
    output = tmp_path / "total.docx"
    doc.save(output)
    assert [row.cells[0].text for row in Document(output).tables[0].rows] == ["Item", "A", "B", "C", "TOTAL"]
    assert Document(output).tables[0].rows[-1].cells[1].text == "60"


@pytest.mark.parametrize("semantic", [False, True])
def test_null_cells_are_empty_and_zero_is_preserved(tmp_path, semantic):
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Amount"
    info = TableInfo(0, ["Item", "Amount"], [1, 2], "Items")
    rows = [{"Item": "A", "Amount": None}, {"Item": "B", "Amount": 0}]
    columns = {0: {0: "Item", 1: "Amount"}} if semantic else None
    fill_document(doc, [], {}, {"items": rows}, [info], table_key_mapping={0: "items"}, column_mappings=columns)
    output = tmp_path / "null.docx"
    doc.save(output)
    amounts = [row.cells[1].text for row in Document(output).tables[0].rows[1:]]
    assert amounts == ["", "0"]


@pytest.mark.parametrize("semantic", [False, True])
def test_null_replaces_a_table_cell_marker_with_empty_text(semantic):
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Amount"
    table.cell(1, 1).text = "______"
    columns = {0: {0: "Item", 1: "Amount"}} if semantic else None
    fill_document(
        doc,
        [],
        {},
        {"items": [{"Item": "A", "Amount": None}]},
        [TableInfo(0, ["Item", "Amount"], [1], "Items")],
        table_key_mapping={0: "items"},
        column_mappings=columns,
    )
    assert table.cell(1, 1).text == ""


def test_data_region_stops_before_totals_and_trailing_spacers():
    doc = Document()
    table = doc.add_table(rows=5, cols=2)
    table.cell(0, 0).text, table.cell(0, 1).text = "Item", "Amount"
    table.cell(3, 0).text = "TOTAL"
    _, _, tables = get_all_placeholders(doc)
    assert tables[0].data_row_indices == [1, 2]
    rows = [{"Item": "A"}, {"Item": "B"}, {"Item": "C"}]
    fill_document(doc, [], {}, {"items": rows}, tables, table_key_mapping={0: "items"})
    assert [row.cells[0].text for row in table.rows] == ["Item", "A", "B", "C", "TOTAL", ""]


def test_table_with_single_blank_data_row():
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text, table.cell(0, 1).text = "Item", "Amount"

    _, _, tables = get_all_placeholders(doc)

    assert len(tables) == 1
    assert tables[0].data_row_indices == [1]


def test_marker_data_row_is_filled_under_multirow_header(tmp_path):
    doc = Document()
    table = doc.add_table(rows=5, cols=2)
    table.cell(0, 0).merge(table.cell(0, 1)).text = "Items sold"
    table.cell(1, 0).text, table.cell(1, 1).text = "Item", "Amount"
    table.cell(2, 0).text, table.cell(2, 1).text = "______", "......"
    table.cell(3, 0).merge(table.cell(3, 1)).text = "Notes"
    _, fields, tables = get_all_placeholders(doc)
    assert len(tables) == 1
    assert tables[0].headers == ["Item", "Amount"]
    assert tables[0].data_row_indices == [2]
    rows = [{"Item": "A", "Amount": None}, {"Item": "B", "Amount": 0}]
    fill_document(doc, fields, {}, {"items": rows}, tables, table_key_mapping={0: "items"})

    output = tmp_path / "markers.docx"
    doc.save(output)
    saved = Document(output).tables[0]
    assert [tuple(cell.text for cell in row.cells) for row in saved.rows] == [
        ("Items sold", "Items sold"),
        ("Item", "Amount"),
        ("A", ""),
        ("B", "0"),
        ("Notes", "Notes"),
        ("", ""),
    ]


@pytest.mark.parametrize(
    "rows,expected",
    [
        (["A", "B"], [["A", ""], ["B", ""]]),
        ([{"Item": "A", "Note": "FIRST NOTE"}, {"Item": "B"}], [["A", "FIRST NOTE"], ["B", ""]]),
    ],
)
def test_array_owned_markers_do_not_receive_scalar_values(tmp_path, rows, expected):
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text, table.cell(0, 1).text = "Item", "Note"
    table.cell(1, 0).text = table.cell(1, 1).text = "______"
    table.cell(2, 0).text = "TOTAL"
    _, fields, tables = get_all_placeholders(doc)
    scalar = {_placeholder_id(field): {"value": "SCALAR VALUE", "used_keys": ["note"]} for field in fields}
    fill_document(doc, fields, scalar, {"items": rows, "note": "SCALAR VALUE"}, tables, table_key_mapping={0: "items"})
    target = tmp_path / "markers.docx"
    doc.save(target)
    actual = [[cell.text for cell in row.cells] for row in Document(target).tables[0].rows]
    assert actual == [["Item", "Note"], *expected, ["TOTAL", ""]]


def test_unmatched_marker_table_keeps_scalar_mapping():
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text, table.cell(0, 1).text = "Name", "Note"
    table.cell(1, 0).text = "______"
    _, fields, tables = get_all_placeholders(doc)
    scalar = {_placeholder_id(fields[0]): {"value": "DEMO Alex", "used_keys": ["name"]}}
    fill_document(doc, fields, scalar, {"name": "DEMO Alex"}, tables)
    assert table.cell(1, 0).text == "DEMO Alex"


def test_array_marker_cleanup_preserves_unmapped_drawing(tmp_path):
    import base64
    from io import BytesIO

    from docx.shared import Inches

    png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aWZ0AAAAASUVORK5CYII="
    )
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text, table.cell(0, 1).text = "Item", "Symbol"
    table.cell(1, 0).text = "______"
    table.cell(1, 1).paragraphs[0].add_run().add_picture(BytesIO(png), width=Inches(0.1))
    table.cell(2, 0).text = "TOTAL"
    _, fields, tables = get_all_placeholders(doc)
    fill_document(
        doc,
        fields,
        {},
        {"items": [{"Item": "A"}]},
        tables,
        table_key_mapping={0: "items"},
        column_mappings={0: {0: "Item"}},
    )
    target = tmp_path / "drawing.docx"
    doc.save(target)
    reopened = Document(target)
    assert len(reopened.inline_shapes) == 1
    assert reopened.tables[0].cell(1, 0).text == "A"
