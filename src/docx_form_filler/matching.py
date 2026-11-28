"""Offline fallback heuristics for matching tables and columns to JSON keys.

Used only when the model provides no mapping. Matching is diacritic-insensitive
and rewards shared words and shared 4-letter prefixes, which copes with Romanian
inflection (``ofertant`` / ``ofertantului``) but can be fooled by ambiguous labels.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Mapping, Sequence
from typing import Any

from docx_form_filler.analyzer import TableInfo

_PUNCTUATION_RE = re.compile(r"[^\w\s]")
_PREFIX_LENGTH = 4


def normalize_text(text: str) -> str:
    """Lowercase, strip diacritics and punctuation."""
    decomposed = unicodedata.normalize("NFKD", text.lower())
    without_marks = "".join(char for char in decomposed if not unicodedata.combining(char))
    return _PUNCTUATION_RE.sub("", without_marks).strip()


def _header_words(header: str) -> set[str]:
    return set(normalize_text(header.replace("\n", " ")).split())


def _key_words(key: str) -> set[str]:
    return set(normalize_text(key.replace("_", " ")).split())


def _similarity(left: set[str], right: set[str]) -> int:
    """3 points per shared word, 2 per distinct pair sharing a 4-letter prefix."""
    score = 3 * len(left & right)
    for a in left:
        for b in right:
            if (
                a != b
                and min(len(a), len(b)) >= _PREFIX_LENGTH
                and (a.startswith(b[:_PREFIX_LENGTH]) or b.startswith(a[:_PREFIX_LENGTH]))
            ):
                score += 2
    return score


def match_table_to_array(table: TableInfo, arrays: Mapping[str, list[Any]]) -> str | None:
    """Best-scoring array key for a table, from its context, headers and row structure."""
    context = normalize_text(table.context)
    headers = " ".join(normalize_text(header.replace("\n", " ")) for header in table.headers)
    header_words = [words for words in map(_header_words, table.headers) if words]

    best_key, best_score = None, 0
    for key, rows in arrays.items():
        score = 0
        for word in normalize_text(key.replace("-", " ").replace("_", " ")).split():
            if len(word) > 3:
                score += 2 * (word in context) + 3 * (word in headers)
        # Object keys that resemble column headers are the most reliable signal.
        if rows and isinstance(rows[0], dict):
            for data_key in rows[0]:
                key_words = _key_words(data_key)
                score += sum(_similarity(key_words, words) for words in header_words)
        if score > best_score:
            best_key, best_score = key, score
    return best_key


def match_columns_by_header(headers: Sequence[str], data_keys: Sequence[str]) -> dict[int, str]:
    """Greedy one-to-one assignment of object keys to columns, highest score first."""
    normalized_headers = [normalize_text(header.strip().replace("\n", " ")) for header in headers]
    normalized_keys = {key: normalize_text(key.replace("_", " ")) for key in data_keys}
    candidates = []
    for column, header in enumerate(normalized_headers):
        if not header:
            continue
        for key, normalized_key in normalized_keys.items():
            score = _similarity(set(normalized_key.split()), set(header.split()))
            if normalized_key in header or header in normalized_key:
                score += 2
            if score > 0:
                candidates.append((score, column, key))

    candidates.sort(key=lambda candidate: -candidate[0])
    assigned: dict[str, str] = {}  # header text -> key; identical headers share a key
    used: set[str] = set()
    for _, column, key in candidates:
        header = headers[column]
        if header not in assigned and key not in used:
            assigned[header] = key
            used.add(key)
    return {column: assigned[header] for column, header in enumerate(headers) if header in assigned}
