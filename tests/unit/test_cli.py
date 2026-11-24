import json
import os
import sys
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document

from main import main, run_pipeline


@pytest.mark.parametrize("payload", [[], None, "value", 5])
def test_requires_json_object(tmp_path, monkeypatch, capsys, payload):
    docx = tmp_path / "form.docx"
    Document().save(docx)
    data = tmp_path / "data.json"
    data.write_text(json.dumps(payload))
    monkeypatch.setattr(sys, "argv", ["main.py", str(docx), str(data)])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 1
    assert "JSON root must be an object" in capsys.readouterr().out
    assert not (tmp_path / "resources").exists()


def test_help_does_not_require_input_files(monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["main.py", "--help"])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 0
    assert "DOCX" in capsys.readouterr().out


def test_invalid_json_has_actionable_error(tmp_path, monkeypatch, capsys):
    docx = tmp_path / "form.docx"
    Document().save(docx)
    data = tmp_path / "data.json"
    data.write_text("{")
    monkeypatch.setattr(sys, "argv", ["main.py", str(docx), str(data)])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 1
    assert "Invalid JSON" in capsys.readouterr().out


def test_corrupt_docx_has_actionable_error(tmp_path, monkeypatch, capsys):
    template = tmp_path / "corrupt.docx"
    template.write_text("not a word archive")
    data = tmp_path / "data.json"
    data.write_text("{}")
    monkeypatch.setattr(sys, "argv", ["main.py", str(template), str(data)])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 1
    assert "Cannot read DOCX" in capsys.readouterr().out


def test_zip_without_word_manifest_has_actionable_error(tmp_path, monkeypatch, capsys):
    template = tmp_path / "archive.docx"
    with ZipFile(template, "w") as archive:
        archive.writestr("readme.txt", "Not a Word document")
    data = tmp_path / "data.json"
    data.write_text("{}")
    monkeypatch.setattr(sys, "argv", ["main.py", str(template), str(data)])

    with pytest.raises(SystemExit) as error:
        main()

    assert error.value.code == 1
    assert "Cannot read DOCX" in capsys.readouterr().out


def test_output_cannot_overwrite_template(tmp_path, monkeypatch, capsys):
    template = tmp_path / "form.docx"
    Document().save(template)
    original = template.read_bytes()
    data = tmp_path / "data.json"
    data.write_text("{}")
    monkeypatch.setattr(sys, "argv", ["main.py", str(template), str(data), "--output", str(template)])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 1
    assert "overwrite the input" in capsys.readouterr().out
    assert template.read_bytes() == original


def test_provider_error_does_not_expose_details(tmp_path, monkeypatch, capsys):
    from openai import OpenAIError

    template = tmp_path / "form.docx"
    doc = Document()
    doc.add_paragraph("Name: ______")
    doc.save(template)
    data = tmp_path / "data.json"
    data.write_text('{"name":"DEMO Alex"}')

    def unavailable(*args, **kwargs):
        raise OpenAIError("sensitive provider details")

    monkeypatch.setattr("main.map_fields_with_llm", unavailable)
    monkeypatch.setattr(sys, "argv", ["main.py", str(template), str(data)])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 1
    output = capsys.readouterr().out
    assert "Model request failed" in output
    assert "sensitive provider details" not in output


def test_hardlinked_output_cannot_overwrite_template(tmp_path, monkeypatch, capsys):
    source, target = tmp_path / "form.docx", tmp_path / "alias.docx"
    Document().save(source)
    os.link(source, target)
    original = source.read_bytes()
    data = tmp_path / "data.json"
    data.write_text("{}")
    monkeypatch.setattr(sys, "argv", ["main.py", str(source), str(data), "--output", str(target)])
    with pytest.raises(SystemExit) as error:
        main()
    assert error.value.code == 1
    assert "overwrite the input" in capsys.readouterr().out
    assert source.read_bytes() == original


def test_failed_save_preserves_existing_output_and_cleans_up(tmp_path, monkeypatch):
    template = tmp_path / "form.docx"
    doc = Document()
    doc.save(template)
    data = tmp_path / "data.json"
    data.write_text("{}")
    destination = tmp_path / "generated"
    destination.mkdir()
    output = destination / "filled.docx"
    doc.save(output)
    original = output.read_bytes()
    monkeypatch.setattr("main.map_fields_with_llm", lambda *args: {})

    def fail_after_writing(self, target):
        Path(target).write_bytes(b"partial DOCX")
        raise OSError("save failed")

    monkeypatch.setattr(type(doc), "save", fail_after_writing)
    with pytest.raises(OSError, match="save failed"):
        run_pipeline(template, data, output)

    assert output.read_bytes() == original
    assert list(destination.iterdir()) == [output]
