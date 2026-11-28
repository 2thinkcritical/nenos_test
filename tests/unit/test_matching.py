"""Offline fallback heuristics used when the model returns no table or column mapping."""

from docx_form_filler.analyzer import TableInfo
from docx_form_filler.matching import match_columns_by_header, match_table_to_array, normalize_text


def test_normalize_text_removes_diacritics_and_punctuation():
    assert normalize_text("Ofertantului (Ţară)!") == "ofertantului tara"


def test_table_matches_array_whose_object_keys_resemble_headers():
    table = TableInfo(0, ["Denumire serviciu", "Preţ (lei)"], [1], "Lista serviciilor")
    arrays = {
        "persons": [{"name": "Alex"}],
        "services": [{"denumire": "Consulting", "pret": "500"}],
    }
    assert match_table_to_array(table, arrays) == "services"


def test_table_without_any_signal_has_no_match():
    table = TableInfo(0, ["Alpha"], [1], "")
    assert match_table_to_array(table, {"items": [{"zzz": 1}]}) is None


def test_columns_are_assigned_one_to_one_by_inflected_prefix():
    headers = ["Nr.", "Denumirea ofertantului", "Valoarea ofertei"]
    assert match_columns_by_header(headers, ["ofertant", "valoare"]) == {1: "ofertant", 2: "valoare"}


def test_a_key_is_never_assigned_to_two_columns():
    assert match_columns_by_header(["Amount", "Amount total"], ["amount"]) == {0: "amount"}
