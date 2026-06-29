from dataclasses import dataclass, field


@dataclass(slots=True)
class TextSpan:
    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    href: str | None = None
    type: str = "text"


@dataclass(slots=True)
class ParagraphNode:
    inlines: list[TextSpan]
    type: str = "paragraph"


@dataclass(slots=True)
class HeadingNode:
    level: int
    inlines: list[TextSpan]
    type: str = "heading"


@dataclass(slots=True)
class QuoteNode:
    inlines: list[TextSpan]
    type: str = "quote"


@dataclass(slots=True)
class HorizontalRuleNode:
    type: str = "horizontal_rule"


@dataclass(slots=True)
class CodeBlockNode:
    content: str
    language: str | None = None
    type: str = "code_block"


@dataclass(slots=True)
class ImageBlockNode:
    raw_path: str
    alt_text: str = ""
    title: str | None = None
    type: str = "image"


@dataclass(slots=True)
class TableNode:
    headers: list[list[TextSpan]]
    rows: list[list[list[TextSpan]]] = field(default_factory=list)
    type: str = "table"


@dataclass(slots=True)
class ListItemNode:
    inlines: list[TextSpan]
    task: bool = False
    checked: bool | None = None
    level: int = 0
    kind: str | None = None
    sequence_start: int | None = None


@dataclass(slots=True)
class ListNode:
    kind: str
    start: int = 1
    items: list[ListItemNode] = field(default_factory=list)
    type: str = "list"
