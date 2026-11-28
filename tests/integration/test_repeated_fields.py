"""Distinct blanks in the same run or cell must not overwrite one another."""

import pytest
from docx import Document

from docx_form_filler.analyzer import analyze_document, placeholder_id
from docx_form_filler.filler import fill_document


@pytest.mark.parametrize("in_table", [False, True])
def test_repeated_blanks_have_independent_mappings(in_table):
    doc = Document()
    paragraph = doc.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if in_table else doc.add_paragraph()
    paragraph.add_run("Name: ______ Address: ______")
    fields, _ = analyze_document(doc)
    assert len(fields) == 2
    assert len({placeholder_id(field) for field in fields}) == 2
    mapping = {
        placeholder_id(fields[0]): {"value": None, "used_keys": []},
        placeholder_id(fields[1]): {"value": "Demo Street", "used_keys": ["address"]},
    }
    fill_document(doc, fields, mapping, {"address": "Demo Street"})
    assert paragraph.text == "Name: ______ Address: Demo Street"


@pytest.mark.parametrize("in_table", [False, True])
def test_only_the_selected_checkbox_is_checked(in_table):
    doc = Document()
    paragraph = doc.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if in_table else doc.add_paragraph()
    paragraph.add_run("|_| First |_| Second")
    fields, _ = analyze_document(doc)
    assert len(fields) == 2
    mapping = {
        placeholder_id(fields[0]): {"value": "UNCHECKED", "used_keys": []},
        placeholder_id(fields[1]): {"value": "CHECKED", "used_keys": []},
    }
    fill_document(doc, fields, mapping, {})
    assert paragraph.text == "|_| First |X| Second"
