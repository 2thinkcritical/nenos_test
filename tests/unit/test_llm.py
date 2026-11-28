"""Parsing of model replies at the application boundary."""

import pytest

from docx_form_filler.llm import load_prompt, parse_json_objects


@pytest.mark.parametrize(
    "content",
    ['[{"id": "a"}]', '```json\n[{"id": "a"}]\n```', '```\n[{"id": "a"}, 3, "x"]\n```'],
)
def test_object_items_are_extracted_from_plain_or_fenced_arrays(content):
    assert parse_json_objects(content) == [{"id": "a"}]


@pytest.mark.parametrize("content", ["", "not json", '{"id": "a"}', "null", "```json\n{]\n```"])
def test_anything_but_an_array_yields_no_items(content):
    assert parse_json_objects(content) == []


@pytest.mark.parametrize("name", ["fields", "tables", "columns"])
def test_prompts_ship_with_the_package(name):
    assert "JSON array" in load_prompt(name)
