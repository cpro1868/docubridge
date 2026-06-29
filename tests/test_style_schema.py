from pathlib import Path

import pytest
from pydantic import ValidationError

from docubridge.core.style_schema import load_style_profile


FIXTURE_PATH = Path(__file__).resolve().parent / "fixtures" / "style.yaml"


def test_load_style_profile_reads_elements_and_defaults() -> None:
    profile = load_style_profile(FIXTURE_PATH)

    assert profile.defaults["font_name"] == "Times New Roman"
    assert profile.elements["heading1"]["font_size"] == 18
    assert profile.elements["paragraph"]["based_on"] == "Normal"


def test_load_style_profile_applies_override_values() -> None:
    profile = load_style_profile(
        FIXTURE_PATH,
        overrides={"document.toc.depth": "4"},
    )

    assert profile.document["toc"]["depth"] == 4


def test_load_style_profile_rejects_missing_override_path() -> None:
    with pytest.raises(KeyError):
        load_style_profile(
            FIXTURE_PATH,
            overrides={"document.missing.depth": "4"},
        )


def test_load_style_profile_rejects_scalar_override_path() -> None:
    with pytest.raises(TypeError):
        load_style_profile(
            FIXTURE_PATH,
            overrides={"defaults.font_name.size": "4"},
        )


def test_load_style_profile_rejects_mapping_leaf_override() -> None:
    with pytest.raises(TypeError):
        load_style_profile(
            FIXTURE_PATH,
            overrides={"document.toc": "4"},
        )


def test_load_style_profile_rejects_unknown_leaf_override_key() -> None:
    with pytest.raises(KeyError):
        load_style_profile(
            FIXTURE_PATH,
            overrides={"document.toc.dept": "4"},
        )


def test_load_style_profile_rejects_unknown_top_level_keys() -> None:
    with pytest.raises(ValidationError):
        load_style_profile(
            FIXTURE_PATH,
            overrides={"unexpected": "value"},
        )
