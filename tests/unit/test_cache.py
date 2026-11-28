"""Regressions for cache invalidation when document semantics change."""

from dataclasses import replace

import pytest

from docx_form_filler.analyzer import PlaceholderInfo, TableInfo, placeholder_id
from docx_form_filler.mapper import load_cached_mapping, map_tables_to_arrays, save_mapping_to_cache


def placeholder(context):
    return PlaceholderInfo("paragraph", 0, 0, "______", context, "", context + "______")


def test_changed_context_invalidates_field_cache():
    original = placeholder("Name: ")
    save_mapping_to_cache(
        {placeholder_id(original): {"value": "Alex", "used_keys": ["name"]}}, [original], {"name": "Alex"}
    )
    assert load_cached_mapping([original], {"name": "Alex"}) is not None
    assert load_cached_mapping([replace(original, context_before="Address: ")], {"name": "Alex"}) is None


@pytest.mark.parametrize("contents", ["[]", "null", '{"inputs_hash": 3}', "{"])
def test_malformed_cache_is_a_miss(contents):
    from pathlib import Path

    Path(".cache").mkdir()
    Path(".cache/mapping.json").write_text(contents)
    assert load_cached_mapping([placeholder("Name: ")], {"name": "Alex"}) is None


def test_changed_array_schema_invalidates_table_cache(monkeypatch):
    from types import SimpleNamespace

    responses = iter(['[{"table_index": 0, "json_key": "items"}]', "[]"])

    class FakeModel:
        def invoke(self, messages):
            return SimpleNamespace(content=next(responses))

    monkeypatch.setattr("docx_form_filler.llm.ChatOpenAI", lambda **kwargs: FakeModel())
    tables = [TableInfo(0, ["Item"], [1], "Items")]
    assert map_tables_to_arrays(tables, {"items": [{"name": "Apple"}]}) == {0: "items"}
    assert map_tables_to_arrays(tables, {"items": [{"unrelated": "Apple"}]}) == {}


def test_changed_model_invalidates_cache(monkeypatch):
    field = placeholder("Name: ")
    monkeypatch.setenv("OPENAI_MODEL", "model-a")
    save_mapping_to_cache({}, [field], {})
    monkeypatch.setenv("OPENAI_MODEL", "model-b")
    assert load_cached_mapping([field], {}) is None


@pytest.mark.parametrize(("value", "keys"), [("YES", ["check"]), ("CHECKED", ["unknown"])])
def test_invalid_checkbox_cache_is_a_miss(value, keys):

    field = replace(placeholder("Select: "), is_checkbox=True)
    save_mapping_to_cache({placeholder_id(field): {"value": value, "used_keys": keys}}, [field], {"check": True})
    assert load_cached_mapping([field], {"check": True}) is None


def test_changed_table_row_count_invalidates_cache(monkeypatch):
    from types import SimpleNamespace

    responses = iter(['[{"table_index": 0, "json_key": "items"}]', "[]"])

    class FakeModel:
        def invoke(self, messages):
            return SimpleNamespace(content=next(responses))

    monkeypatch.setattr("docx_form_filler.llm.ChatOpenAI", lambda **kwargs: FakeModel())
    table = TableInfo(0, ["Item"], [1], "Items")
    data = {"items": [{"name": "Apple"}]}
    assert map_tables_to_arrays([table], data) == {0: "items"}
    assert map_tables_to_arrays([replace(table, data_row_indices=[1, 2])], data) == {}
