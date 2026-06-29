from __future__ import annotations

from pathlib import Path
from typing import Any

from ruamel.yaml import YAML


def load_yaml(path: Path) -> dict[str, Any]:
    yaml = YAML(typ="safe")
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.load(handle)
    if data is None:
        return {}
    if not isinstance(data, dict):
        raise TypeError(f"Expected mapping at YAML root in {path}")
    return data


def write_yaml(path: Path, data: dict[str, Any], *, pretty: bool = False) -> None:
    yaml = YAML()
    if pretty:
        yaml.default_flow_style = False
        yaml.indent(mapping=2, sequence=4, offset=2)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        yaml.dump(data, handle)
