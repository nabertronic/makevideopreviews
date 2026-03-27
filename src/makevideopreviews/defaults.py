from pathlib import Path
import os


DEFAULT_INTERVAL = 10
DEFAULT_MAX_WIDTH = 480
DEFAULT_QUALITY = 10
DEFAULT_PAGE_WIDTH_INCH = 10.0
DEFAULT_WORKERS = min(4, max(1, os.cpu_count() or 1))
DEFAULT_VIDEO_ROOT = Path.home() / "Desktop"

VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".mkv", ".avi", ".mxf"}
