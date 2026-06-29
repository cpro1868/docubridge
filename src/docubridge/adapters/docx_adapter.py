from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn


def _strip_document_content(document: DocumentType) -> None:
    """Remove body content while preserving styles, numbering and section properties."""
    body = document.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def create_document(template_path: Path | None = None) -> DocumentType:
    if template_path is None:
        return Document()
    document = Document(str(template_path))
    _strip_document_content(document)
    return document
