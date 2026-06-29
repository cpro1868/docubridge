from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field


class RenderRequest(BaseModel):
    input_path: Path
    output_path: Path
    style_path: Path | None = None
    template_path: Path | None = None
    profile_name: str | None = None
    mode: Literal["strict", "lenient"] = "strict"
    output_mode: Literal["human", "quiet", "json"] = "human"
    overwrite: bool = False
    resource_dir: Path | None = None
    dump_ast: bool = False
    features: list[str] = Field(default_factory=list)
    overrides: dict[str, str] = Field(default_factory=dict)
