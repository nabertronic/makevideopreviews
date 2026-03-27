from __future__ import annotations

from fractions import Fraction
import math
import os
import re
import unicodedata
from pathlib import Path
from tempfile import NamedTemporaryFile


def clamp_quality(value: int) -> int:
    return max(2, min(31, value))


def seconds_to_hms(seconds: float) -> str:
    total_seconds = int(round(max(0, seconds)))
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    secs = total_seconds % 60
    return "%02d:%02d:%02d" % (hours, minutes, secs)


def human_bytes(size: int) -> str:
    value = float(size)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if value < 1024.0 or unit == "TB":
            if unit == "B":
                return "%d %s" % (int(value), unit)
            return "%.1f %s" % (value, unit)
        value /= 1024.0
    return "%.1f TB" % value


def safe_output_name(name: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9._-]+", "_", name.strip())
    return normalized or "root"


def parse_frame_rate(raw_value: str) -> float:
    if not raw_value or raw_value in {"0/0", "N/A"}:
        return 0.0
    return float(Fraction(raw_value))


def estimated_frame_count(duration: float, interval: int) -> int:
    if not duration or duration <= 0:
        return 0
    return int(math.floor(duration / max(1, interval))) + 1


def format_resolution(width: int, height: int) -> str:
    if not width or not height:
        return "unknown"
    return "%sx%s" % (width, height)


def atomic_docx_path(target: Path) -> Path:
    tmp_file = NamedTemporaryFile(
        prefix=".%s_" % target.stem,
        suffix=".docx",
        dir=str(target.parent),
        delete=False,
    )
    tmp_path = Path(tmp_file.name)
    tmp_file.close()
    return tmp_path


def can_write_to(path: Path) -> bool:
    try:
        probe_file = NamedTemporaryFile(prefix=".writecheck_", dir=str(path), delete=True)
        probe_file.close()
        return True
    except OSError:
        return False


def unusual_run_hint(total_frames: int, total_bytes: int) -> str:
    if total_frames >= 250000 or total_bytes >= 5 * 1024 * 1024 * 1024:
        return "Large documentary-scale run detected. Expect long extraction times and very large .docx output."
    return ""


def sanitize_user_path(raw_value) -> str:
    text = str(raw_value).strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {"'", '"'}:
        text = text[1:-1].strip()
    return os.path.expanduser(text)


def resolve_existing_dir(raw_value) -> Path:
    cleaned = sanitize_user_path(raw_value)
    if not cleaned:
        return Path(cleaned)

    direct = Path(cleaned)
    if direct.exists() and direct.is_dir():
        return direct.resolve()

    normalized = Path(unicodedata.normalize("NFC", cleaned))
    if normalized.exists() and normalized.is_dir():
        return normalized.resolve()

    normalized = Path(unicodedata.normalize("NFD", cleaned))
    if normalized.exists() and normalized.is_dir():
        return normalized.resolve()

    resolved = _resolve_by_segments(cleaned)
    if resolved is not None and resolved.is_dir():
        return resolved.resolve()

    return direct.expanduser()


def _resolve_by_segments(raw_value: str):
    original = Path(raw_value)
    anchor = original.anchor or os.sep
    current = Path(anchor)
    parts = [part for part in original.parts[1:] if part not in {"", os.sep}]

    if not current.exists():
        return None

    for part in parts:
        if not current.is_dir():
            return None
        next_path = current / part
        if next_path.exists():
            current = next_path
            continue

        match = _find_matching_child(current, part)
        if match is None:
            return None
        current = match

    return current


def _find_matching_child(parent: Path, wanted: str):
    wanted_keys = _path_keys(wanted)
    try:
        children = list(parent.iterdir())
    except OSError:
        return None

    for child in children:
        if child.name == wanted:
            return child

    for child in children:
        child_keys = _path_keys(child.name)
        if child_keys[0] == wanted_keys[0]:
            return child

    for child in children:
        child_keys = _path_keys(child.name)
        if child_keys[1] == wanted_keys[1]:
            return child

    return None


def _path_keys(text: str):
    stripped = sanitize_user_path(text)
    nfc = unicodedata.normalize("NFC", stripped).casefold()
    folded = "".join(ch for ch in unicodedata.normalize("NFKD", stripped).casefold() if not unicodedata.combining(ch))
    simple = re.sub(r"[\s'\"`’“”._()-]+", "", folded)
    return nfc, simple
