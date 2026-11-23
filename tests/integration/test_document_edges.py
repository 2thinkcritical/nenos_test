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
