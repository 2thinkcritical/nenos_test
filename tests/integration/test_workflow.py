import json
import sys
from unittest.mock import MagicMock, patch

import pytest
from docx import Document

from docx_form_filler.analyzer import analyze_document, placeholder_id
from docx_form_filler.cli import main


@pytest.mark.parametrize("serialized", [False, True], ids=["native-array", "encoded-array"])
@patch("docx_form_filler.llm.ChatOpenAI")
def test_end_to_end_workflow(mock_chat_openai, tmp_path, capsys, serialized):
    # 1. Setup Input Files
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
            {"offering": "Consulting", "amount_ron": "500"},
            {"offering": "Development", "amount_ron": "1000"},
        ],
    }
    if serialized:
        data["services"] = json.dumps(data["services"])
    json_path.write_text(json.dumps(data))

    # 2. Mock LLM Responses
    mock_llm = MagicMock()
    mock_chat_openai.return_value = mock_llm

    # We need to handle multiple calls to invoke:
    # 1. map_fields_with_llm (batches)
    # 2. map_tables_to_arrays
    # 3. map_table_columns_with_llm

    def side_effect(messages):
        # We need to find the IDs from the HumanMessage which contains the placeholders
        content_for_ids = ""
        for m in messages:
            content_for_ids += m.content + "\n"

        if "DOCX placeholders to fill" in messages[1].content:
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
        if messages[1].content.startswith("DOCX tables:"):
            # This is map_tables_to_arrays
            return MagicMock(content=json.dumps([{"table_index": 0, "json_key": "services"}]))
        if messages[1].content.startswith("Table column headers"):
            # This is map_table_columns_with_llm
            return MagicMock(
                content=json.dumps(
                    [{"column_index": 0, "dict_key": "offering"}, {"column_index": 1, "dict_key": "amount_ron"}]
                )
            )
        return MagicMock(content="[]")

    mock_llm.invoke.side_effect = side_effect

    # 3. Run Main
    with patch.object(sys, "argv", ["docx-fill", str(docx_path), str(json_path)]):
        main()

    output_file = tmp_path / "test_form.filled.docx"
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

    assert table.cell(2, 0).text == "Development"
    assert table.cell(2, 1).text == "1000"


@patch("docx_form_filler.llm.ChatOpenAI")
def test_custom_output_path(mock_model, tmp_path, monkeypatch):
    doc = Document()
    doc.add_paragraph("Name: ______")
    source = tmp_path / "source.docx"
    doc.save(source)
    data = tmp_path / "input.json"
    data.write_text('{"name":"DEMO Alex"}')
    target = tmp_path / "nested" / "result.docx"
    fields, _ = analyze_document(doc)
    pid = placeholder_id(fields[0])
    mock_model.return_value.invoke.return_value.content = json.dumps(
        [{"id": pid, "value": "DEMO Alex", "used_json_keys": ["name"]}]
    )
    monkeypatch.setattr(sys, "argv", ["docx-fill", str(source), str(data), "--output", str(target)])
    main()
    assert Document(target).paragraphs[0].text == "Name: DEMO Alex"
    assert not (tmp_path / "resources").exists()


@patch("docx_form_filler.llm.ChatOpenAI")
def test_sparse_array_includes_keys_from_later_rows(mock_model, tmp_path):
    from types import SimpleNamespace

    from docx_form_filler.pipeline import run_pipeline

    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text, table.cell(0, 1).text = "Item", "Amount"
    source, data, target = tmp_path / "source.docx", tmp_path / "data.json", tmp_path / "out.docx"
    doc.save(source)
    data.write_text(json.dumps({"items": [{"item": "A"}, {"item": "B", "amount": 0}]}))
    replies = iter(
        [
            [{"table_index": 0, "json_key": "items"}],
            [{"column_index": 0, "dict_key": "item"}, {"column_index": 1, "dict_key": "amount"}],
        ]
    )
    mock_model.return_value.invoke.side_effect = lambda messages: SimpleNamespace(content=json.dumps(next(replies)))
    run_pipeline(source, data, target)
    reopened = Document(target).tables[0]
    assert [[cell.text for cell in row.cells] for row in reopened.rows[1:]] == [["A", ""], ["B", "0"]]
