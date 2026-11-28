"""Plain-JSON decision caches with atomic writes."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any


class JsonCache:
    """A JSON object stored in one file. Unreadable or malformed content reads as empty."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def read(self) -> dict[str, Any]:
        try:
            value = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}
        return value if isinstance(value, dict) else {}

    def write(self, data: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(
            "w", encoding="utf-8", dir=self.path.parent, prefix=f".{self.path.name}-", delete=False
        ) as temporary:
            json.dump(data, temporary, indent=2, ensure_ascii=False)
        Path(temporary.name).replace(self.path)


def fingerprint(payload: object) -> str:
    """Short, order-independent hash of JSON-serialisable cache inputs."""
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()[:16]
