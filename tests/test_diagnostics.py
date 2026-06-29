from pathlib import Path

from docubridge import __version__
from docubridge.application.models import RenderRequest
from docubridge.core.diagnostics import Diagnostic, Severity


def test_package_version_is_exposed():
    assert __version__ == "0.1.0"


def test_diagnostic_to_dict_contains_expected_fields() -> None:
    diagnostic = Diagnostic(
        severity=Severity.WARNING,
        code="STYLE_UNKNOWN_FIELD",
        message="Unknown style field",
        location="style.yaml:12",
        hint="Remove the field or rename it",
    )

    assert diagnostic.to_dict() == {
        "severity": "warning",
        "code": "STYLE_UNKNOWN_FIELD",
        "message": "Unknown style field",
        "location": "style.yaml:12",
        "hint": "Remove the field or rename it",
    }


def test_render_request_normalizes_paths() -> None:
    request = RenderRequest(
        input_path="tests/fixtures/sample.md",
        output_path="build/output.docx",
        style_path="tests/fixtures/style.yaml",
        template_path=None,
        profile_name="academic",
        mode="strict",
        output_mode="human",
        overwrite=False,
        resource_dir="assets",
        dump_ast=False,
        features=["toc"],
        overrides={"document.toc.depth": "3"},
    )

    assert isinstance(request.input_path, Path)
    assert isinstance(request.output_path, Path)
    assert isinstance(request.style_path, Path)
    assert request.input_path.as_posix().endswith("tests/fixtures/sample.md")
    assert request.output_path.as_posix().endswith("build/output.docx")
    assert request.resource_dir.as_posix().endswith("assets")
