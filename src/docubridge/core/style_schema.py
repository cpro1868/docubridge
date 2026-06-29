from __future__ import annotations

from pathlib import Path
from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from docubridge.adapters.yaml_adapter import load_yaml


class StyleProfile(BaseModel):
    model_config = ConfigDict(extra="forbid")

    meta: dict[str, Any] = Field(default_factory=dict)
    defaults: dict[str, Any] = Field(default_factory=dict)
    elements: dict[str, dict[str, Any]] = Field(default_factory=dict)
    multilevel_list: dict[str, Any] = Field(default_factory=dict)
    document: dict[str, Any] = Field(default_factory=dict)
    assets: dict[str, Any] = Field(default_factory=dict)
    compat: dict[str, Any] = Field(default_factory=dict)


def coerce_scalar(value: str) -> Any:
    text = value.strip()
    lowered = text.lower()
    if lowered in {"null", "~"}:
        return None
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return value


def apply_dotted_override(target: dict[str, Any], dotted_path: str, raw_value: str) -> None:
    parts = dotted_path.split(".")
    if not parts or any(not part for part in parts):
        raise ValueError(f"Invalid override path: {dotted_path}")

    cursor: dict[str, Any] = target
    for key in parts[:-1]:
        if key not in cursor:
            raise KeyError(f"Missing override path segment: {key} in {dotted_path}")
        next_value = cursor[key]
        if not isinstance(next_value, dict):
            raise TypeError(f"Override path segment is not a mapping: {key} in {dotted_path}")
        cursor = next_value

    leaf_key = parts[-1]
    if len(parts) > 1 and leaf_key not in cursor:
        raise KeyError(f"Missing override leaf key: {leaf_key} in {dotted_path}")
    if isinstance(cursor.get(leaf_key), dict):
        raise TypeError(f"Override target is a mapping: {leaf_key} in {dotted_path}")
    cursor[leaf_key] = coerce_scalar(raw_value)


def load_style_profile(path: Path, overrides: dict[str, str] | None = None) -> StyleProfile:
    data = load_yaml(path)
    for dotted_path, raw_value in (overrides or {}).items():
        apply_dotted_override(data, dotted_path, raw_value)
    return StyleProfile.model_validate(data)
