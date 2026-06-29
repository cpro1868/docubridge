from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class RunStyleIntent:
    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    href: str | None = None


@dataclass(slots=True)
class NumberingIntent:
    numbering_role: str
    level: int = 0
    continue_sequence: bool = True
    start_at: int | None = None
    preferred_template_style: str | None = None


@dataclass(slots=True)
class ParagraphLayoutIntent:
    element_name: str
    runs: list[RunStyleIntent] = field(default_factory=list)
    resolved_style_name: str = "Normal"
    resolved_properties: dict[str, Any] = field(default_factory=dict)
    prefix_text: str = ""
    numbering: NumberingIntent | None = None
