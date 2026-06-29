from pathlib import Path

from markdown_it import MarkdownIt
from markdown_it.token import Token


def build_markdown_parser() -> MarkdownIt:
    parser = MarkdownIt("gfm-like")
    parser.enable("table")
    return parser


def load_markdown_tokens(path: Path) -> list[Token]:
    return build_markdown_parser().parse(path.read_text(encoding="utf-8"))
