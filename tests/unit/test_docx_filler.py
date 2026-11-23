import pytest
from docx import Document

from docx_filler import PlaceholderInfo, TableInfo, fill_document


@pytest.fixture
def sample_placeholders():
    return [
        PlaceholderInfo(
            location_type="paragraph",
            paragraph_index=0,
            run_index=0,
            placeholder_pattern="______",
            context_before="Name: ",
            context_after="",
            full_paragraph_text="Name: ______",
            section_index=0,
            header_footer_type="body",
        ),
        PlaceholderInfo(
            location_type="paragraph",
            paragraph_index=1,
            run_index=0,
            placeholder_pattern="|_|",
            context_before="Check: ",
            context_after="",
            full_paragraph_text="Check: |_|",
            section_index=0,
            header_footer_type="body",
            is_checkbox=True,
            checkbox_label="Check me",
        ),
    ]


@pytest.fixture
def sample_mapping():
    return {
        "body_para_0_run0_______": {"value": "John Doe", "used_keys": ["name"]},
        "body_para_1_run0_|_|": {"value": "CHECKED", "used_keys": ["check"]},
    }


def test_fill_document_paragraphs(tmp_path, sample_placeholders, sample_mapping):
    doc = Document()
    doc.add_paragraph("Name: ______")
    doc.add_paragraph("Check: |_|")

    fill_document(doc, sample_placeholders, sample_mapping, {})

    assert "John Doe" in doc.paragraphs[0].text
    assert "______" not in doc.paragraphs[0].text
    assert "|X|" in doc.paragraphs[1].text
    assert "|_|" not in doc.paragraphs[1].text


def test_fill_data_tables(tmp_path):
    doc = Document()
    # Create table with header and empty data rows
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Cost"

    data_tables = [
        TableInfo(
            table_index=0,
            headers=["Item", "Cost"],
            data_row_indices=[1, 2],
            context="List of items",
            section_index=0,
            header_footer_type="body",
        )
    ]

    json_data = {"items": [{"Item": "Apple", "Cost": "10"}, {"Item": "Banana", "Cost": "20"}]}

    table_key_mapping = {0: "items"}

    fill_document(doc, [], {}, json_data, data_tables, table_key_mapping)

    assert "Apple" == doc.tables[0].cell(1, 0).text
    assert "10" == doc.tables[0].cell(1, 1).text
    assert "Banana" == doc.tables[0].cell(2, 0).text
    assert "20" == doc.tables[0].cell(2, 1).text
