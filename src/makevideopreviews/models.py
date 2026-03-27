from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import List, Optional, Union


class CommandMode(str, Enum):
    GENERATE = "generate"
    ESTIMATE = "estimate"


class OverwriteMode(str, Enum):
    SKIP = "skip"
    OVERWRITE = "overwrite"
    SKIP_EXISTING = "skip_existing"


class PreviewScope(str, Enum):
    PROJECT = "project"
    SUBFOLDER = "subfolder"


class WorkerMode(str, Enum):
    AUTO = "auto"


@dataclass
class AppConfig:
    command: CommandMode
    root: Path
    preview_scope: PreviewScope
    interval: int
    max_width: int
    quality: int
    page_width_inch: float
    workers: Union[int, WorkerMode]
    overwrite_mode: OverwriteMode
    debug: bool = False


@dataclass
class VideoProbe:
    path: Path
    duration: Optional[float]
    width: Optional[int]
    height: Optional[int]
    fps: Optional[float]
    estimated_frames: int
    problem: Optional[str] = None


@dataclass
class FolderJob:
    folder: Path
    videos: List[Path]
    output_path: Path
    probes: List[VideoProbe] = field(default_factory=list)
    estimated_bytes: int = 0
    warnings: List[str] = field(default_factory=list)


@dataclass
class ExtractionResult:
    video_path: Path
    frame_paths: List[Path]
    timestamps: List[float]
    temp_dir: Optional[Path]
    error: Optional[str] = None


@dataclass
class RunSummary:
    folders_total: int = 0
    folders_written: int = 0
    folders_skipped: int = 0
    folders_failed: int = 0
    videos_total: int = 0
    estimated_total_bytes: int = 0
    written_paths: List[Path] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
