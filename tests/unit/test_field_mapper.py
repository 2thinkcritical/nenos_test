from unittest.mock import MagicMock, patch

import pytest

from field_mapper import PlaceholderInfo, TableInfo, map_fields_with_llm, map_table_to_json_key_with_llm


@pytest.fixture
def mock_placeholders():
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
def mock_tables():
    return [
        TableInfo(
            table_index=0,
            headers=["Item", "Cost"],
            data_row_indices=[1, 2],
            context="List of items",
            section_index=0,
            header_footer_type="body",
        )
    ]


@patch("field_mapper.ChatOpenAI")
def test_map_fields_with_llm(mock_chat_openai, mock_placeholders):
    # Setup mock response
    mock_llm_instance = MagicMock()
    mock_chat_openai.return_value = mock_llm_instance

    mock_response_content = """
    [
        {"id": "body_para_0_run0_______", "value": "John Doe", "used_json_keys": ["name"]},
        {"id": "body_para_1_run0_|_|", "value": "CHECKED", "used_json_keys": ["check"]}
    ]
    """
    mock_llm_instance.invoke.return_value.content = mock_response_content

    json_data = {"name": "John Doe", "check": True}

    mapping = map_fields_with_llm(mock_placeholders, ["name", "check"], json_data)

    assert len(mapping) == 2
    assert mapping["body_para_0_run0_______"]["value"] == "John Doe"
    assert mapping["body_para_1_run0_|_|"]["value"] == "CHECKED"


@patch("field_mapper.ChatOpenAI")
def test_map_table_to_json_key_with_llm(mock_chat_openai, mock_tables):
    mock_llm_instance = MagicMock()
    mock_chat_openai.return_value = mock_llm_instance

    mock_response_content = """
    [
        {"table_index": 0, "json_key": "items"}
    ]
    """
    mock_llm_instance.invoke.return_value.content = mock_response_content

    array_keys = {"items": [{"item": "Apple", "cost": "10"}]}

    mapping = map_table_to_json_key_with_llm(mock_tables, array_keys)

    assert len(mapping) == 1
    assert mapping[0] == "items"


@pytest.mark.parametrize(
    "response", ["{}", "null", "[{}]", '[{"id":"unknown","value":"invented","used_json_keys":[]}]']
)
@patch("field_mapper.ChatOpenAI")
def test_invalid_field_responses_leave_fields_unfilled(mock_chat_openai, mock_placeholders, response):
    mock_chat_openai.return_value.invoke.return_value.content = response
    mapping = map_fields_with_llm(mock_placeholders, ["name"], {"name": "Alex"})
    assert len(mapping) == 2
    assert all(item["value"] is None for item in mapping.values())


@pytest.mark.parametrize("response", ["null", "{}", "[{}]", '[{"table_index":99,"json_key":"items"}]'])
@patch("field_mapper.ChatOpenAI")
def test_invalid_table_responses_do_not_create_mappings(mock_chat_openai, mock_tables, response):
    mock_chat_openai.return_value.invoke.return_value.content = response
    assert map_table_to_json_key_with_llm(mock_tables, {"items": [{"item": "Apple"}]}) == {}


@pytest.mark.parametrize(
    "item",
    [
        {"id": "body_para_1_run0_|_|", "value": "YES", "used_json_keys": ["check"]},
        {"id": "body_para_0_run0_______", "value": "Alex"},
    ],
)
@patch("field_mapper.ChatOpenAI")
def test_invalid_checkbox_or_missing_source_list_stays_unfilled(mock_chat_openai, mock_placeholders, item):
    import json

    mock_chat_openai.return_value.invoke.return_value.content = json.dumps([item])
    mapping = map_fields_with_llm(mock_placeholders, ["name", "check"], {"name": "Alex", "check": True})
    assert all(value["value"] is None for value in mapping.values())
