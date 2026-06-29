from __future__ import annotations

from pathlib import Path

from docubridge.adapters.markdown_adapter import load_markdown_tokens


def scan_markdown_warnings(path: Path) -> list[str]:
    warning_order = {
        "markdown contains code block content that is not fully rendered yet": 0,
        "markdown contains images that may be degraded or replaced": 1,
        "markdown contains raw HTML blocks that will be degraded": 2,
        "markdown contains raw HTML inline content that will be degraded": 3,
    }
    seen: set[str] = set()

    def add(message: str) -> None:
        seen.add(message)

    for token in load_markdown_tokens(path):
        if token.type in {"fence", "code_block"}:
            add("markdown contains code block content that is not fully rendered yet")
        if token.type == "html_block":
            add("markdown contains raw HTML blocks that will be degraded")
        if token.type == "inline":
            for child in getattr(token, "children", []) or []:
                if child.type == "image":
                    add("markdown contains images that may be degraded or replaced")
                if child.type == "html_inline":
                    add("markdown contains raw HTML inline content that will be degraded")

    return sorted(seen, key=lambda message: warning_order.get(message, len(warning_order)))
