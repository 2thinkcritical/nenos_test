import pytest
import sys
import json
from pathlib import Path
from unittest.mock import patch, MagicMock
from main import main
from docx import Document

@patch("field_mapper.ChatOpenAI")
def test_end_to_end_workflow(mock_chat_openai, tmp_path, capsys):
    # 1. Setup Input Files
    import shutil
    if (Path(".cache")).exists():
        shutil.rmtree(".cache")
        
    docx_path = tmp_path / "test_form.docx"
    json_path = tmp_path / "test_data.json"
    
    # Create DOCX
    doc = Document()
    doc.add_paragraph("Contract Agreement")
    doc.add_paragraph("Date: ......")
    doc.add_paragraph("Provider: ______")
    doc.add_paragraph("Services Table:")
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Service"
    table.cell(0, 1).text = "Price"
    table.cell(1, 0).text = ""
    table.cell(1, 1).text = ""
    table.cell(2, 0).text = ""
    table.cell(2, 1).text = ""
    doc.save(docx_path)
    
    # Create JSON
    data = {
        "date": "2023-10-27",
        "provider": "Acme Corp",
        "services": [
            {"Service": "Consulting", "Price": "500"},
            {"Service": "Development", "Price": "1000"}
        ]
    }
    with open(json_path, "w") as f:
        json.dump(data, f)

    # 2. Mock LLM Responses
    mock_llm = MagicMock()
    mock_chat_openai.return_value = mock_llm
    
    # We need to handle multiple calls to invoke:
    # 1. map_fields_with_llm (batches)
    # 2. map_table_to_json_key_with_llm
    # 3. map_table_columns_with_llm
    
    def side_effect(messages):
        # We need to find the IDs from the HumanMessage which contains the placeholders
        content_for_ids = ""
        for m in messages:
            content_for_ids += m.content + "\n"
            
        if "filling procurement documents" in messages[0].content:
            # This is map_fields_with_llm
            # Parse the prompt to find the actual IDs being requested
            import re
            ids = re.findall(r'"id": "(.*?)"', content_for_ids)
            
            # Construct response based on requested IDs to be robust
            response_data = []
            for pid in ids:
                if "......" in pid:
                    response_data.append({"id": pid, "value": "2023-10-27", "used_json_keys": ["date"]})
                elif "______" in pid:
                    response_data.append({"id": pid, "value": "Acme Corp", "used_json_keys": ["provider"]})
            
            return MagicMock(content=json.dumps(response_data))
        elif "match each table to the correct JSON key" in content:
             # This is map_table_to_json_key_with_llm
             return MagicMock(content=json.dumps([
                 {"table_index": 0, "json_key": "services"}
             ]))
        elif "match each column to the correct dict key" in content:
            # This is map_table_columns_with_llm
            return MagicMock(content=json.dumps([
                {"column_index": 0, "dict_key": "Service"},
                {"column_index": 1, "dict_key": "Price"}
            ]))
        return MagicMock(content="[]")

    mock_llm.invoke.side_effect = side_effect

    # 3. Run Main
    with patch.object(sys, 'argv', ["main.py", str(docx_path), str(json_path)]):
        main()

    # 4. Verify Output
    # main.py saves to "resources" directory relative to current working dir, 
    # but we need to check where it actually saved. 
    # main.py does: output_dir = Path("resources"); output_dir.mkdir(...)
    # It uses relative path.
    
    # We can check stdout to find the output path or look in ./resources
    # But running in tmp_path might be safer if we change cwd.
    # main.py doesn't change cwd. It uses "resources" in CWD.
    
    output_file = Path("resources") / "test_form.filled.docx"
    assert output_file.exists()
    
    # 5. Verify Content
    doc_filled = Document(output_file)
    text = "\n".join([p.text for p in doc_filled.paragraphs])
    assert "2023-10-27" in text
    assert "Acme Corp" in text
    assert "......" not in text
    assert "______" not in text
    
    table = doc_filled.tables[0]
    assert table.cell(1, 0).text == "Consulting"
    assert table.cell(1, 1).text == "500"
