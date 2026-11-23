"""Distinct blanks in the same run or cell must not overwrite one another."""

import pytest
from docx import Document

from docx_analyzer import get_all_placeholders
from docx_filler import fill_document
from field_mapper import _placeholder_id


@pytest.mark.parametrize("in_table", [False, True])
def test_repeated_blanks_have_independent_mappings(in_table):
    doc = Document()
    paragraph = doc.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if in_table else doc.add_paragraph()
    paragraph.add_run("Name: ______ Address: ______")
    _, fields, _ = get_all_placeholders(doc)
    assert len(fields) == 2
    assert len({_placeholder_id(field) for field in fields}) == 2
    mapping = {
        _placeholder_id(fields[0]): {"value": None, "used_keys": []},
        _placeholder_id(fields[1]): {"value": "Demo Street", "used_keys": ["address"]},
    }
    fill_document(doc, fields, mapping, {"address": "Demo Street"})
    assert paragraph.text == "Name: ______ Address: Demo Street"


@pytest.mark.parametrize("in_table", [False, True])
def test_only_the_selected_checkbox_is_checked(in_table):
    doc = Document()
    paragraph = doc.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if in_table else doc.add_paragraph()
    paragraph.add_run("|_| First |_| Second")
    _, fields, _ = get_all_placeholders(doc)
    assert len(fields) == 2
    mapping = {
        _placeholder_id(fields[0]): {"value": "UNCHECKED", "used_keys": []},
        _placeholder_id(fields[1]): {"value": "CHECKED", "used_keys": []},
    }
    fill_document(doc, fields, mapping, {})
    assert paragraph.text == "|_| First |X| Second"
