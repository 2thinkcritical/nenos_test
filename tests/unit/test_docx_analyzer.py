from docx import Document

from docx_analyzer import extract_placeholders_recursive


def test_extract_placeholders_paragraphs(tmp_path):
    # Create a dummy docx
    doc = Document()
    doc.add_paragraph("This is a test document.")
    doc.add_paragraph("Name: ______")
    doc.add_paragraph("Date: ......")
    doc.add_paragraph("Checkbox: |_|")

    # Run extraction
    _, placeholders, tables = extract_placeholders_recursive(doc)

    assert len(placeholders) == 3

    # Verify underscore placeholder
    p1 = placeholders[0]
    assert "______" in p1.placeholder_pattern
    assert p1.location_type == "paragraph"

    # Verify dot placeholder
    p2 = placeholders[1]
    assert "......" in p2.placeholder_pattern

    # Verify checkbox
    p3 = placeholders[2]
    assert "|_|" in p3.placeholder_pattern
    assert p3.is_checkbox is True


def test_extract_placeholders_table(tmp_path):
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Header 1"
    table.cell(0, 1).text = "Header 2"
    table.cell(1, 0).text = "Value: ____"
    table.cell(1, 1).text = "Check: |_|"

    _, placeholders, tables = extract_placeholders_recursive(doc)

    assert len(placeholders) == 2

    p1 = placeholders[0]
    assert p1.location_type == "table"
    assert "____" in p1.placeholder_pattern
    assert p1.table_index == 0
    assert p1.row_index == 1
    assert p1.cell_index == 0

    p2 = placeholders[1]
    assert p2.location_type == "table"
    assert "|_|" in p2.placeholder_pattern
    assert p2.is_checkbox is True


def test_extract_data_tables(tmp_path):
    doc = Document()
    table = doc.add_table(rows=4, cols=2)
    # Header row
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Cost"
    # Empty data rows
    table.cell(1, 0).text = ""
    table.cell(1, 1).text = ""
    table.cell(2, 0).text = ""
    table.cell(2, 1).text = ""
    table.cell(3, 0).text = ""
    table.cell(3, 1).text = ""

    _, placeholders, tables = extract_placeholders_recursive(doc)

    assert len(tables) == 1
    t = tables[0]
    assert t.headers == ["Item", "Cost"]
    assert len(t.data_row_indices) == 3
