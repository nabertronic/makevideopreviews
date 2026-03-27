from __future__ import annotations

from pathlib import Path

from makevideopreviews.defaults import (
    DEFAULT_INTERVAL,
    DEFAULT_QUALITY,
    DEFAULT_VIDEO_ROOT,
)
from makevideopreviews.models import AppConfig, CommandMode, OverwriteMode, PreviewScope, WorkerMode
from makevideopreviews.utils import resolve_existing_dir


def prompt_for_run():
    try:
        import questionary
        from prompt_toolkit.styles import Style
    except ModuleNotFoundError as exc:
        raise RuntimeError("Interactive mode requires questionary: %s" % exc)

    style = Style.from_dict(
        {
            "qmark": "fg:#d97706 bold",
            "question": "bold",
            "answer": "fg:#1d4ed8 bold",
            "pointer": "fg:#d97706 bold",
            "highlighted": "fg:#92400e bold",
            "selected": "fg:#0f766e",
        }
    )

    action = questionary.select(
        "What do you want to do?",
        choices=[
            questionary.Choice("Generate .docx previews", value=CommandMode.GENERATE),
            questionary.Choice("Estimate output size only", value=CommandMode.ESTIMATE),
            questionary.Choice("Run dependency doctor", value="doctor"),
        ],
        default=CommandMode.GENERATE,
        style=style,
    ).ask()
    if action == "doctor":
        return "doctor", DEFAULT_VIDEO_ROOT

    root = questionary.path(
        "VIDEO_ROOT",
        default=str(DEFAULT_VIDEO_ROOT),
        only_directories=True,
        style=style,
    ).ask()
    preview_scope = questionary.select(
        "Preview layout",
        choices=[
            questionary.Choice("One preview file for the whole project", value=PreviewScope.PROJECT),
            questionary.Choice("One preview file per video subfolder", value=PreviewScope.SUBFOLDER),
        ],
        default=PreviewScope.PROJECT,
        style=style,
    ).ask()
    interval = int(questionary.text("Interval in seconds", default=str(DEFAULT_INTERVAL), style=style).ask())
    quality = questionary.select(
        "Preview image quality",
        choices=[
            questionary.Choice("Balanced (recommended)", value=10),
            questionary.Choice("Higher detail, larger file", value=6),
            questionary.Choice("Smaller file, more compression", value=14),
            questionary.Choice("Custom numeric ffmpeg quality", value="custom"),
        ],
        default=DEFAULT_QUALITY,
        style=style,
    ).ask()
    if quality == "custom":
        quality = int(
            questionary.text(
                "Custom image quality (2 = highest detail, 31 = smallest file)",
                default=str(DEFAULT_QUALITY),
                style=style,
            ).ask()
        )
    overwrite = questionary.select(
        "Existing preview documents",
        choices=[
            questionary.Choice("Skip existing files", value=OverwriteMode.SKIP),
            questionary.Choice("Overwrite existing files", value=OverwriteMode.OVERWRITE),
            questionary.Choice("Quietly skip existing files", value=OverwriteMode.SKIP_EXISTING),
        ],
        default=OverwriteMode.SKIP,
        style=style,
    ).ask()
    debug = bool(questionary.confirm("Enable debug output?", default=False, style=style).ask())

    config = AppConfig(
        command=action,
        root=resolve_existing_dir(root),
        preview_scope=preview_scope,
        interval=interval,
        max_width=480,
        quality=quality,
        page_width_inch=10.0,
        workers=WorkerMode.AUTO,
        overwrite_mode=overwrite,
        debug=debug,
    )
    return "run", config
