"""
Input parser for TaskBridge.

Prefix rules:
  /c  → Calendar only
  /t  → Todoist only
  (none) → both Calendar and Todoist
"""

from dataclasses import dataclass
from typing import Literal

Target = Literal["both", "calendar", "todoist"]


@dataclass
class ParsedInput:
    target: Target
    text: str


def parse_input(raw: str) -> ParsedInput:
    """Parse raw user input and return target destination + clean text."""
    stripped = raw.strip()

    if stripped.startswith("/c ") or stripped == "/c":
        return ParsedInput(target="calendar", text=stripped[3:].strip())

    if stripped.startswith("/t ") or stripped == "/t":
        return ParsedInput(target="todoist", text=stripped[3:].strip())

    return ParsedInput(target="both", text=stripped)
