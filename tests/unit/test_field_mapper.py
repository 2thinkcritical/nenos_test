import pytest
from unittest.mock import MagicMock, patch
from field_mapper import map_fields_with_llm, map_table_to_json_key_with_llm, PlaceholderInfo, TableInfo

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
            header_footer_type="body"
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
            checkbox_label="Check me"
        )
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
            header_footer_type="body"
        )
    ]

@patch("field_mapper.ChatOpenAI")
def test_map_fields_with_llm(mock_chat_openai, mock_placeholders):
    # Setup mock response
    mock_llm_instance = MagicMock()
    mock_chat_openai.return_value = mock_llm_instance
    
    mock_response_content = """
    [
        {"id": "body_para_0_run0______", "value": "John Doe", "used_json_keys": ["name"]},
        {"id": "body_para_1_run0_|_|", "value": "CHECKED", "used_json_keys": ["check"]}
    ]
    """
    mock_llm_instance.invoke.return_value.content = mock_response_content
    
    json_data = {"name": "John Doe", "check": True}
    
    mapping = map_fields_with_llm(mock_placeholders, ["name", "check"], json_data)
    
    assert len(mapping) == 2
    assert mapping["body_para_0_run0______"]["value"] == "John Doe"
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
