from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class PipelineConfig:
    input_path: Path
    explicit_output_dir: str | None = None


@dataclass
class FileRecord:
    source_path: Path
    output_paths: list[Path] = field(default_factory=list)
    status: str = "pending"  # pending | ok | skip | fail
    error: str | None = None


@dataclass
class PipelineState:
    config: PipelineConfig
    base_output_dir: Path | None = None
    files: list[Path] = field(default_factory=list)
    suffixes: set[str] = field(default_factory=set)
    records: list[FileRecord] = field(default_factory=list)
    succeeded: int = 0
    skipped: int = 0
    failures: list[tuple[Path, str]] = field(default_factory=list)
    exit_code: int = 0
