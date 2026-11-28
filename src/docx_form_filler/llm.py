"""The single boundary between the application and the chat model.

Everything model-specific lives here: configuration, invocation and parsing of
the reply into untrusted JSON objects. Callers validate the objects they receive.
"""

from __future__ import annotations

import json
import os
from functools import cache
from importlib.resources import files
from typing import Any

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI as ChatOpenAI  # re-exported: the seam tests replace

DEFAULT_MODEL = "gpt-5.1"
REQUEST_TIMEOUT_SECONDS = 120
MAX_RETRIES = 3


def model_name() -> str:
    return os.environ.get("OPENAI_MODEL", DEFAULT_MODEL)


@cache
def load_prompt(name: str) -> str:
    """Read a system prompt shipped in ``docx_form_filler/prompts``."""
    return files(__package__).joinpath("prompts", f"{name}.md").read_text(encoding="utf-8")


def request_json_objects(prompt_name: str, user_message: str) -> list[dict[str, Any]]:
    """Ask the model for a JSON array and return its object items.

    Transport errors propagate (the client already retries transient failures);
    a reply that is not a JSON array yields an empty list.
    """
    model = ChatOpenAI(
        model=model_name(),
        temperature=0,
        timeout=REQUEST_TIMEOUT_SECONDS,
        max_retries=MAX_RETRIES,
    )
    response = model.invoke([SystemMessage(content=load_prompt(prompt_name)), HumanMessage(content=user_message)])
    return parse_json_objects(response.content)


def parse_json_objects(content: object) -> list[dict[str, Any]]:
    """Extract the object items of a JSON array, tolerating a fenced code block."""
    text = str(content).strip()
    if text.startswith("```"):
        text = text.partition("\n")[2].removesuffix("```").strip()
    try:
        value = json.loads(text)
    except ValueError:
        return []
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def to_json(value: object) -> str:
    """Pretty JSON for prompts, keeping non-ASCII text readable."""
    return json.dumps(value, indent=2, ensure_ascii=False)
