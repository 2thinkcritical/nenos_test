import json
import sys

import pytest
from docx import Document

from main import main


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
