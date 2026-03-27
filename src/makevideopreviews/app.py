from __future__ import annotations

from pathlib import Path
import importlib.util
import sys
import tempfile

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text
from rich.theme import Theme
from rich import box

from makevideopreviews.defaults import (
    DEFAULT_INTERVAL,
    DEFAULT_MAX_WIDTH,
    DEFAULT_PAGE_WIDTH_INCH,
    DEFAULT_QUALITY,
    DEFAULT_VIDEO_ROOT,
    DEFAULT_WORKERS,
)
from makevideopreviews.discovery import discover_jobs
from makevideopreviews.errors import DependencyError, MakeVideoPreviewsError, ValidationError
from makevideopreviews.interactive import prompt_for_run
from makevideopreviews.media import (
    cleanup_extraction_results,
    ensure_ffmpeg_tools,
    estimate_jobs,
    extract_job_frames,
    determine_worker_count,
    populate_jobs_with_probes,
)
from makevideopreviews.models import AppConfig, CommandMode, OverwriteMode, PreviewScope, RunSummary, WorkerMode
from makevideopreviews.utils import can_write_to, human_bytes, resolve_existing_dir, unusual_run_hint


APP_THEME = Theme(
    {
        "accent": "bold #d97706",
        "accent.soft": "#b45309",
        "ok": "bold #0f766e",
        "warn": "bold #b45309",
        "error": "bold #b91c1c",
        "muted": "#6b7280",
        "info": "#0369a1",
        "path": "bold #1d4ed8",
    }
)

console = Console(theme=APP_THEME)
app = typer.Typer(
    add_completion=False,
    help="Create compact Word previews for large video folders.",
    invoke_without_command=True,
    no_args_is_help=False,
)


@app.callback(invoke_without_command=True)
def main(ctx: typer.Context):
    if ctx.invoked_subcommand is not None:
        return
    if "--help" in sys.argv or "-h" in sys.argv:
        return
    if not sys.stdin.isatty():
        console.print("No command given. Use --help for CLI usage.", style="warn")
        raise typer.Exit(code=1)

    _print_banner()
    selection, payload = prompt_for_run()
    if selection == "doctor":
        raise typer.Exit(code=run_doctor(payload))
    validate_config(payload)
    raise typer.Exit(code=run_pipeline(payload))


@app.command()
def generate(
    root: str = typer.Option(str(DEFAULT_VIDEO_ROOT), "--root"),
    scope: PreviewScope = typer.Option(PreviewScope.PROJECT, "--scope"),
    interval: int = typer.Option(DEFAULT_INTERVAL, "--interval"),
    max_width: int = typer.Option(DEFAULT_MAX_WIDTH, "--max-width"),
    quality: int = typer.Option(DEFAULT_QUALITY, "--quality"),
    page_width_inch: float = typer.Option(DEFAULT_PAGE_WIDTH_INCH, "--page-width-inch"),
    workers: int = typer.Option(None, "--workers"),
    overwrite: bool = typer.Option(False, "--overwrite"),
    skip_existing: bool = typer.Option(False, "--skip-existing"),
    debug: bool = typer.Option(False, "--debug"),
):
    config = build_config(
        command=CommandMode.GENERATE,
        root=root,
        preview_scope=scope,
        interval=interval,
        max_width=max_width,
        quality=quality,
        page_width_inch=page_width_inch,
        workers=workers,
        overwrite=overwrite,
        skip_existing=skip_existing,
        debug=debug,
    )
    raise typer.Exit(code=run_pipeline(config))


@app.command()
def estimate(
    root: str = typer.Option(str(DEFAULT_VIDEO_ROOT), "--root"),
    scope: PreviewScope = typer.Option(PreviewScope.PROJECT, "--scope"),
    interval: int = typer.Option(DEFAULT_INTERVAL, "--interval"),
    max_width: int = typer.Option(DEFAULT_MAX_WIDTH, "--max-width"),
    quality: int = typer.Option(DEFAULT_QUALITY, "--quality"),
    page_width_inch: float = typer.Option(DEFAULT_PAGE_WIDTH_INCH, "--page-width-inch"),
    workers: int = typer.Option(None, "--workers"),
    debug: bool = typer.Option(False, "--debug"),
):
    config = build_config(
        command=CommandMode.ESTIMATE,
        root=root,
        preview_scope=scope,
        interval=interval,
        max_width=max_width,
        quality=quality,
        page_width_inch=page_width_inch,
        workers=workers,
        overwrite=False,
        skip_existing=False,
        debug=debug,
    )
    raise typer.Exit(code=run_pipeline(config))


@app.command()
def doctor(
    root: str = typer.Option(str(DEFAULT_VIDEO_ROOT), "--root"),
):
    raise typer.Exit(code=run_doctor(root))


def run(argv=None):
    if argv is not None:
        sys.argv = [sys.argv[0]] + list(argv)
    app()
    return 0


def build_config(command, root, preview_scope, interval, max_width, quality, page_width_inch, workers, overwrite, skip_existing, debug):
    overwrite_mode = OverwriteMode.SKIP
    if overwrite:
        overwrite_mode = OverwriteMode.OVERWRITE
    elif skip_existing:
        overwrite_mode = OverwriteMode.SKIP_EXISTING

    config = AppConfig(
        command=command,
        root=resolve_existing_dir(root),
        preview_scope=preview_scope,
        interval=interval,
        max_width=max_width,
        quality=quality,
        page_width_inch=page_width_inch,
        workers=workers if workers is not None else WorkerMode.AUTO,
        overwrite_mode=overwrite_mode,
        debug=debug,
    )
    validate_config(config)
    return config


def validate_config(config: AppConfig):
    if config.interval < 1:
        raise typer.BadParameter("Interval must be >= 1.")
    if config.max_width < 64:
        raise typer.BadParameter("Max width must be >= 64.")
    if config.quality < 2 or config.quality > 31:
        raise typer.BadParameter("Quality must be between 2 and 31.")
    if config.page_width_inch <= 0:
        raise typer.BadParameter("Page width must be > 0.")
    if isinstance(config.workers, int) and config.workers < 1:
        raise typer.BadParameter("Workers must be >= 1.")
    if not config.root.exists() or not config.root.is_dir():
        raise typer.BadParameter("Root path must exist and be a directory.")


def run_pipeline(config: AppConfig) -> int:
    try:
        ensure_ffmpeg_tools()
        jobs, discovery_warnings = discover_jobs(config.root, config.preview_scope)
        if not jobs:
            console.print(
                Panel(
                    "No folders with video files found under\n[%s]%s[/]." % ("path", config.root),
                    title="Nothing To Do",
                    border_style="warn",
                    box=box.ROUNDED,
                )
            )
            return 0

        summary = RunSummary(folders_total=len(jobs), warnings=list(discovery_warnings))
        total_videos, total_duration = populate_jobs_with_probes(jobs, config)
        config.workers = determine_worker_count(config, jobs)
        summary.videos_total = total_videos
        summary.estimated_total_bytes = estimate_jobs(jobs, config)
        _print_run_context(config)
        _print_preflight(jobs, summary, total_duration)

        if config.command == CommandMode.ESTIMATE:
            _print_estimate_summary(jobs, summary)
            return 0

        for job in jobs:
            if job.output_path.exists():
                if config.overwrite_mode == OverwriteMode.SKIP:
                    console.print(
                        Panel.fit(
                            "Skipping existing file\n[%s]%s[/]" % ("path", job.output_path),
                            title="Existing Output",
                            border_style="warn",
                            box=box.ROUNDED,
                        )
                    )
                    summary.folders_skipped += 1
                    continue
                if config.overwrite_mode == OverwriteMode.SKIP_EXISTING:
                    summary.folders_skipped += 1
                    continue

            results = {}
            try:
                console.print(
                    Panel.fit(
                        "[accent]%s[/]\n[path]%s[/]" % (job.output_path.name, job.folder),
                        title="Generating Folder Preview",
                        border_style="accent",
                        box=box.ROUNDED,
                    )
                )
                results = extract_job_frames(job, config)
                from makevideopreviews.docx_render import write_docx

                output = write_docx(job, results, config)
                summary.folders_written += 1
                summary.written_paths.append(output)
                console.print("Wrote [path]%s[/]" % output, style="ok")
            except KeyboardInterrupt:
                cleanup_extraction_results(results.values())
                raise
            except Exception as exc:
                summary.folders_failed += 1
                summary.errors.append("%s: %s" % (job.folder, exc))
                console.print("Failed [path]%s[/]: %s" % (job.folder, exc), style="error")
            finally:
                cleanup_extraction_results(results.values())

        _print_run_summary(summary)
        return 0 if not summary.errors else 1
    except KeyboardInterrupt:
        console.print("\nCancelled by user.", style="warn")
        return 130
    except (DependencyError, ValidationError, MakeVideoPreviewsError) as exc:
        console.print(Panel(str(exc), title="Error", border_style="error", box=box.ROUNDED))
        return 1


def run_doctor(root: Path) -> int:
    checks = []

    for binary in ("ffmpeg", "ffprobe"):
        status = "ok" if _binary_exists(binary) else "missing"
        checks.append((binary, status, _binary_hint(binary)))

    for package_name in ("typer", "rich", "questionary", "docx", "PIL"):
        available = importlib.util.find_spec(package_name) is not None
        status = "ok" if available else "missing"
        checks.append((package_name, status, _package_hint(package_name, available)))

    target_root = resolve_existing_dir(root)
    if target_root.exists():
        writable = can_write_to(target_root)
        checks.append(("write access", "ok" if writable else "missing", str(target_root)))
    else:
        checks.append(("write access", "missing", "%s does not exist" % target_root))
    checks.append(("temp dir", "ok" if can_write_to(Path(tempfile.gettempdir())) else "missing", tempfile.gettempdir()))

    table = Table(title="Doctor Report", box=box.SIMPLE_HEAVY, header_style="accent")
    table.add_column("Check")
    table.add_column("Status")
    table.add_column("Details")
    failed = False
    for name, status, details in checks:
        if status != "ok":
            failed = True
        table.add_row(name, _status_text(status), details)

    console.print(table)
    if failed:
        console.print(
            Panel(
                "Help and estimate may still work, but generation needs all required dependencies.",
                title="Doctor Found Issues",
                border_style="warn",
                box=box.ROUNDED,
            )
        )
        return 1
    console.print(Panel("Doctor found no issues.", title="Doctor", border_style="ok", box=box.ROUNDED))
    return 0


def _print_preflight(jobs, summary, total_duration):
    total_frames = sum(probe.estimated_frames for job in jobs for probe in job.probes)
    table = Table(title="Preflight Summary", box=box.SIMPLE_HEAVY, header_style="accent")
    table.add_column("Folders")
    table.add_column("Videos")
    table.add_column("Duration")
    table.add_column("Frames")
    table.add_column("Estimated Size")
    table.add_row(
        str(summary.folders_total),
        str(summary.videos_total),
        _format_duration(total_duration),
        str(total_frames),
        human_bytes(summary.estimated_total_bytes),
    )
    console.print(table)

    folder_table = Table(title="Folders", box=box.MINIMAL_DOUBLE_HEAD, header_style="accent.soft")
    folder_table.add_column("Folder")
    folder_table.add_column("Videos")
    folder_table.add_column("Estimated Size")
    for job in jobs:
        folder_table.add_row(str(job.folder), str(len(job.videos)), human_bytes(job.estimated_bytes))
    console.print(folder_table)

    hint = unusual_run_hint(total_frames, summary.estimated_total_bytes)
    if hint:
        console.print(Panel(hint, title="Load Hint", border_style="warn", box=box.ROUNDED))

    warnings = list(summary.warnings)
    for job in jobs:
        warnings.extend(job.warnings)
    if warnings:
        console.print(Panel.fit("Potential issues detected during scanning.", title="Warnings", border_style="warn", box=box.ROUNDED))
        for warning in warnings[:20]:
            console.print("- %s" % warning, style="warn")
        if len(warnings) > 20:
            console.print("- ... %d more" % (len(warnings) - 20), style="muted")


def _print_estimate_summary(jobs, summary):
    console.print(Panel("Estimate-only mode complete.", title="Estimate", border_style="ok", box=box.ROUNDED))
    for job in jobs:
        console.print("[path]%s[/] -> [accent]%s[/]" % (job.folder, human_bytes(job.estimated_bytes)))
    console.print(
        Panel.fit(
            "Estimated total: [accent]%s[/]" % human_bytes(summary.estimated_total_bytes),
            border_style="info",
            box=box.ROUNDED,
        )
    )


def _print_run_summary(summary: RunSummary):
    table = Table(title="Run Summary", box=box.SIMPLE_HEAVY, header_style="accent")
    table.add_column("Metric")
    table.add_column("Value")
    table.add_row("Folders total", str(summary.folders_total))
    table.add_row("Folders written", "[ok]%s[/]" % summary.folders_written)
    table.add_row("Folders skipped", "[warn]%s[/]" % summary.folders_skipped)
    table.add_row("Folders failed", "[error]%s[/]" % summary.folders_failed)
    table.add_row("Videos total", str(summary.videos_total))
    table.add_row("Estimated total size", human_bytes(summary.estimated_total_bytes))
    console.print(table)

    if summary.written_paths:
        console.print(Panel.fit("Written files", border_style="ok", box=box.ROUNDED))
        for path in summary.written_paths:
            console.print("- [path]%s[/]" % path)
    if summary.errors:
        console.print(Panel.fit("Errors", border_style="error", box=box.ROUNDED))
        for error in summary.errors:
            console.print("- %s" % error, style="error")


def _print_banner():
    title = Text("Make Video Previews", style="accent")
    subtitle = Text("Local .docx previews for large documentary footage trees", style="muted")
    console.print(
        Panel(
            Text.assemble(title, "\n", subtitle),
            border_style="accent",
            box=box.DOUBLE,
            padding=(1, 2),
        )
    )


def _print_run_context(config: AppConfig):
    mode_label = "Generate .docx" if config.command == CommandMode.GENERATE else "Estimate only"
    quality_label = "quality level %s" % config.quality
    text = (
        "[accent]Mode:[/] %s\n"
        "[accent]Root:[/] [path]%s[/]\n"
        "[accent]Scope:[/] %s\n"
        "[accent]Interval:[/] %ss\n"
        "[accent]Image detail:[/] %spx source frames\n"
        "[accent]Image quality:[/] %s\n"
        "[accent]Workers:[/] %s"
    ) % (
        mode_label,
        config.root,
        "One file per project" if config.preview_scope == PreviewScope.PROJECT else "One file per subfolder",
        config.interval,
        config.max_width,
        quality_label,
        config.workers,
    )
    console.print(Panel(text, title="Run Context", border_style="info", box=box.ROUNDED))


def _status_text(status):
    styles = {
        "ok": "ok",
        "missing": "error",
    }
    return "[%s]%s[/]" % (styles.get(status, "muted"), status)


def _binary_exists(name):
    from shutil import which

    return which(name) is not None


def _binary_hint(name):
    if _binary_exists(name):
        from shutil import which

        return which(name) or ""
    return "Install %s and ensure it is on PATH." % name


def _package_hint(name, available):
    if available:
        return "Available in the current Python environment."
    mapping = {
        "docx": "Install with `python3 -m pip install python-docx`.",
        "PIL": "Install with `python3 -m pip install Pillow`.",
    }
    return mapping.get(name, "Install with `python3 -m pip install %s`." % name)


def _format_duration(total_seconds):
    total_seconds = int(round(total_seconds))
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return "%02d:%02d:%02d" % (hours, minutes, seconds)
