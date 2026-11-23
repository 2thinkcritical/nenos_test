import pytest
from docx import Document


@pytest.fixture
def sample_docx_path(tmp_path):
    """Creates a sample DOCX file for testing."""
    doc = Document()
    doc.add_paragraph("Hello world")
    doc.add_paragraph("Name: ______")
    doc.add_paragraph("Date: ......")

    # Add a table
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Cost"
    table.cell(1, 0).text = "Apple"
    table.cell(1, 1).text = "____"

    path = tmp_path / "sample.docx"
    doc.save(path)
    return path


@pytest.fixture
def sample_json_data():
    """Returns sample JSON data for testing."""
    return {
        "name": "John Doe",
        "date": "2023-10-27",
        "items": [{"item": "Banana", "cost": "10"}, {"item": "Orange", "cost": "20"}],
    }


@pytest.fixture(autouse=True)
def isolated_workspace(tmp_path, monkeypatch):
    """Keep caches, logs and generated documents outside the checkout."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("OPENAI_API_KEY", "test-key-not-a-credential")
