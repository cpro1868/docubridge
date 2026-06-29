from __future__ import annotations

from pathlib import Path


def resolve_image_path(raw_path: str, base_dir: Path) -> Path:
    if not raw_path or not raw_path.strip():
        raise ValueError("Image path cannot be blank")

    path = Path(raw_path).expanduser()
    if path.is_absolute():
        raise ValueError("Image path resolves outside base_dir")

    resolved_base_dir = Path(base_dir).expanduser().resolve()
    resolved_path = (resolved_base_dir / path).resolve()
    if not resolved_path.is_relative_to(resolved_base_dir):
        raise ValueError("Image path resolves outside base_dir")
    return resolved_path
