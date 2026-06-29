from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from docubridge.core.diagnostics import Diagnostic, Severity
from docubridge.core.docx_ingest import parse_docx_file
from docubridge.core.pptx_ingest import parse_pptx_file
from docubridge.core.xlsx_ingest import parse_xlsx_file


@dataclass(slots=True)
class ParseResult:
    success: bool
    output_path: Path
    diagnostics: list[dict] = field(default_factory=list)


def _failure_result(output_path: Path, code: str, message: str) -> ParseResult:
    diagnostic = Diagnostic(
        severity=Severity.ERROR,
        code=code,
        message=message,
    )
    return ParseResult(
        success=False,
        output_path=output_path,
        diagnostics=[diagnostic.to_dict()],
    )


def run_parse(input_path: Path, output_path: Path) -> ParseResult:
    try:
        suffix = input_path.suffix.lower()
        if suffix == ".docx":
            media_dir = output_path.parent / "assets"
            markdown = parse_docx_file(input_path, media_dir=media_dir)
        elif suffix == ".xlsx":
            markdown = parse_xlsx_file(input_path)
        elif suffix == ".pptx":
            media_dir = output_path.parent / "assets"
            markdown = parse_pptx_file(input_path, media_dir=media_dir)
        else:
            return _failure_result(
                output_path,
                "INPUT_FORMAT_UNSUPPORTED",
                f"Unsupported input format: {input_path.suffix or input_path.name}",
            )
    except FileNotFoundError as exc:
        message = exc.strerror or str(exc)
        if exc.filename:
            message = f"{message}: {exc.filename}"
        return _failure_result(output_path, "INPUT_FILE_NOT_FOUND", message)
    except OSError as exc:
        message = exc.strerror or str(exc)
        if exc.filename:
            message = f"{message}: {exc.filename}"
        return _failure_result(output_path, "INPUT_IO_ERROR", message)

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(markdown, encoding="utf-8")
    except OSError as exc:
        message = exc.strerror or str(exc)
        if exc.filename:
            message = f"{message}: {exc.filename}"
        return _failure_result(output_path, "OUTPUT_IO_ERROR", message)

    return ParseResult(success=True, output_path=output_path)
