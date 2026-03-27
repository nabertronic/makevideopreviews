from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
import time
from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

from rich.progress import BarColumn, Progress, SpinnerColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn, track

from makevideopreviews.errors import DependencyError, FrameExtractionError
from makevideopreviews.models import AppConfig, ExtractionResult, FolderJob, VideoProbe
from makevideopreviews.utils import clamp_quality, estimated_frame_count, parse_frame_rate, safe_output_name

FFPROBE_TIMEOUT_SECONDS = 30
SINGLE_FRAME_TIMEOUT_SECONDS = 60


def ensure_ffmpeg_tools() -> None:
    missing = [tool for tool in ("ffmpeg", "ffprobe") if shutil.which(tool) is None]
    if missing:
        raise DependencyError("Missing external tools: %s" % ", ".join(missing))


def determine_worker_count(config: AppConfig, jobs: List[FolderJob]) -> int:
    if isinstance(config.workers, int):
        return max(1, config.workers)

    total_videos = sum(len(job.videos) for job in jobs)
    cpu_workers = min(4, max(1, shutil.os.cpu_count() or 1))
    cloud_hint = str(config.root).lower()
    is_cloud_path = any(marker in cloud_hint for marker in ("onedrive", "icloud", "dropbox", "googledrive", "cloudstorage"))

    if total_videos <= 2:
        return 1
    if is_cloud_path:
        return 1 if total_videos <= 6 else 2
    if total_videos <= 4:
        return min(2, cpu_workers)
    return cpu_workers


def probe_video(path: Path, interval: int, cache: Dict[str, VideoProbe], debug: bool = False) -> VideoProbe:
    cache_key = str(path.resolve())
    if cache_key in cache:
        return cache[cache_key]

    cmd = [
        "ffprobe",
        "-v",
        "error",
        "-select_streams",
        "v:0",
        "-show_entries",
        "stream=width,height,avg_frame_rate,r_frame_rate:format=duration",
        "-of",
        "json",
        str(path),
    ]
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            timeout=FFPROBE_TIMEOUT_SECONDS,
        )
        payload = json.loads(result.stdout or "{}")
        streams = payload.get("streams") or []
        fmt = payload.get("format") or {}
        stream = streams[0] if streams else {}
        duration = _safe_float(fmt.get("duration"))
        fps = parse_frame_rate(stream.get("avg_frame_rate") or stream.get("r_frame_rate") or "")
        probe = VideoProbe(
            path=path,
            duration=duration,
            width=_safe_int(stream.get("width")),
            height=_safe_int(stream.get("height")),
            fps=fps if fps > 0 else None,
            estimated_frames=estimated_frame_count(duration or 0.0, interval),
            problem=None,
        )
        if duration is None:
            probe.problem = "Could not determine duration."
        cache[cache_key] = probe
        return probe
    except subprocess.TimeoutExpired:
        if debug:
            print("[DEBUG] ffprobe timed out for %s" % path)
        probe = VideoProbe(
            path=path,
            duration=None,
            width=None,
            height=None,
            fps=None,
            estimated_frames=0,
            problem="ffprobe timed out after %ss" % FFPROBE_TIMEOUT_SECONDS,
        )
        cache[cache_key] = probe
        return probe
    except Exception as exc:
        if debug:
            print("[DEBUG] ffprobe failed for %s: %s" % (path, exc))
        probe = VideoProbe(
            path=path,
            duration=None,
            width=None,
            height=None,
            fps=None,
            estimated_frames=0,
            problem="ffprobe failed",
        )
        cache[cache_key] = probe
        return probe


def populate_jobs_with_probes(jobs: List[FolderJob], config: AppConfig) -> Tuple[int, float]:
    cache = {}
    total_videos = sum(len(job.videos) for job in jobs)
    total_duration = 0.0

    for job in track(jobs, description="Probing videos", transient=True):
        probes = []
        for video_path in job.videos:
            probe = probe_video(video_path, config.interval, cache, debug=config.debug)
            probes.append(probe)
            if probe.duration:
                total_duration += probe.duration
            elif probe.problem:
                job.warnings.append("%s: %s" % (video_path.name, probe.problem))
        job.probes = probes

    return total_videos, total_duration


def estimate_jobs(jobs: List[FolderJob], config: AppConfig) -> int:
    total_bytes = 0
    total_videos = sum(len(job.probes) for job in jobs)
    if total_videos <= 0:
        return 0

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("{task.completed}/{task.total} videos"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        transient=True,
    ) as progress:
        task_id = progress.add_task("Estimating output size", total=total_videos)

        def advance():
            progress.advance(task_id)

        for job in jobs:
            job.estimated_bytes = estimate_job_bytes(job, config, progress_callback=advance)
            total_bytes += job.estimated_bytes
    return total_bytes


def estimate_job_bytes(job: FolderJob, config: AppConfig, progress_callback=None) -> int:
    tmp_root = Path(tempfile.mkdtemp(prefix="mvp_estimate_"))
    total_bytes = 0
    try:
        for probe in job.probes:
            if not probe.duration or probe.duration <= 0:
                if progress_callback:
                    progress_callback()
                continue
            sample_positions = [0.0, max(0.0, probe.duration * 0.5), max(0.0, probe.duration * 0.95)]
            sample_sizes = []
            for index, position in enumerate(sample_positions):
                sample_path = tmp_root / ("%s_%02d.jpg" % (safe_output_name(probe.path.name), index))
                if extract_single_frame(probe.path, position, sample_path, config.max_width, config.quality):
                    sample_sizes.append(sample_path.stat().st_size)
            average = int(sum(sample_sizes) / len(sample_sizes)) if sample_sizes else 120000
            total_bytes += int(probe.estimated_frames * average * 1.05)
            if progress_callback:
                progress_callback()
    finally:
        shutil.rmtree(str(tmp_root), ignore_errors=True)
    return total_bytes


def extract_single_frame(src: Path, t_seconds: float, out_path: Path, max_width: int, quality: int) -> bool:
    vf = "scale='min(%s,iw)':-2" % max_width
    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-loglevel",
        "error",
        "-ss",
        str(t_seconds),
        "-i",
        str(src),
        "-frames:v",
        "1",
        "-q:v",
        str(clamp_quality(quality)),
        "-vf",
        vf,
        "-pix_fmt",
        "yuvj420p",
        "-y",
        str(out_path),
    ]
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            timeout=SINGLE_FRAME_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return False
    return result.returncode == 0 and out_path.exists() and out_path.stat().st_size > 0


def extract_job_frames(job: FolderJob, config: AppConfig) -> Dict[Path, ExtractionResult]:
    temp_root = Path(tempfile.mkdtemp(prefix="mvp_frames_"))
    results = {}
    expected_total = sum(max(0, probe.estimated_frames) for probe in job.probes)
    if expected_total <= 0:
        expected_total = max(1, len(job.probes))

    with ProcessPoolExecutor(max_workers=config.workers) as executor:
        futures = {
            executor.submit(
                _extract_video_frames_worker,
                str(probe.path),
                probe.duration or 0.0,
                config.interval,
                config.max_width,
                config.quality,
                str(temp_root),
                config.debug,
            ): probe.path
            for probe in job.probes
        }

        completed = set()
        progress_count = 0
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total} frames"),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            transient=True,
        ) as progress:
            task_id = progress.add_task("Extracting frames", total=expected_total)
            pending = set(futures.keys())

            while pending:
                counted = _count_extracted_frames(temp_root)
                if counted > progress_count:
                    progress.update(task_id, completed=min(counted, expected_total))
                    progress_count = counted

                done, pending = wait(pending, timeout=0.5, return_when=FIRST_COMPLETED)
                for future in done:
                    source_path = futures[future]
                    completed.add(future)
                    try:
                        result = future.result()
                        results[Path(result.video_path)] = result
                    except Exception as exc:
                        results[source_path] = ExtractionResult(
                            video_path=source_path,
                            frame_paths=[],
                            timestamps=[],
                            temp_dir=None,
                            error=str(exc),
                        )
                if not done:
                    time.sleep(0.2)

            final_count = _count_extracted_frames(temp_root)
            progress.update(task_id, completed=min(final_count, expected_total))

    return results


def cleanup_extraction_results(results: Iterable[ExtractionResult]) -> None:
    cleaned = set()
    parents = set()
    for result in results:
        if result.temp_dir and result.temp_dir not in cleaned:
            parents.add(result.temp_dir.parent)
            shutil.rmtree(str(result.temp_dir), ignore_errors=True)
            cleaned.add(result.temp_dir)
    for parent in parents:
        try:
            parent.rmdir()
        except OSError:
            pass


def _extract_video_frames_worker(
    video_path: str,
    duration: float,
    interval: int,
    max_width: int,
    quality: int,
    temp_root: str,
    debug: bool,
) -> ExtractionResult:
    video = Path(video_path)
    work_dir = Path(temp_root) / ("%s_frames" % safe_output_name(video.name))
    work_dir.mkdir(parents=True, exist_ok=True)
    output_pattern = work_dir / "frame_%06d.jpg"
    vf = "fps=1/%s,scale='min(%s,iw)':-2" % (interval, max_width)
    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-loglevel",
        "error",
        "-i",
        str(video),
        "-vf",
        vf,
        "-q:v",
        str(clamp_quality(quality)),
        "-pix_fmt",
        "yuvj420p",
        "-start_number",
        "0",
        "-y",
        str(output_pattern),
    ]
    timeout_seconds = _batch_extract_timeout(duration)
    try:
        process = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired:
        raise FrameExtractionError("ffmpeg timed out after %ss" % timeout_seconds)
    if process.returncode != 0:
        raise FrameExtractionError(process.stderr.strip() or "ffmpeg extraction failed")

    frame_paths = sorted(work_dir.glob("frame_*.jpg"))
    if not frame_paths:
        fallback_path = work_dir / "frame_000000.jpg"
        if extract_single_frame(video, 0.0, fallback_path, max_width, quality):
            frame_paths = [fallback_path]
        else:
            raise FrameExtractionError("ffmpeg produced no preview frames")

    timestamps = []
    for index, _frame in enumerate(frame_paths):
        timestamp = float(index * interval)
        if duration > 0:
            timestamp = min(timestamp, duration)
        timestamps.append(timestamp)

    return ExtractionResult(
        video_path=video,
        frame_paths=frame_paths,
        timestamps=timestamps,
        temp_dir=work_dir,
        error=None,
    )


def _safe_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _safe_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _count_extracted_frames(temp_root: Path) -> int:
    total = 0
    try:
        for frame_path in temp_root.rglob("frame_*.jpg"):
            if frame_path.is_file():
                total += 1
    except OSError:
        return total
    return total


def _batch_extract_timeout(duration: float) -> int:
    safe_duration = max(0.0, duration or 0.0)
    # Allow long documentary clips while still preventing endless hangs.
    return int(max(180, min(7200, 120 + safe_duration * 1.5)))
