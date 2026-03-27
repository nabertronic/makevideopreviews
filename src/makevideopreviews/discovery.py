from __future__ import annotations

import os
from pathlib import Path
from typing import List, Tuple

from makevideopreviews.defaults import VIDEO_EXTS
from makevideopreviews.models import FolderJob, PreviewScope
from makevideopreviews.utils import safe_output_name


def discover_jobs(root: Path, preview_scope: PreviewScope) -> Tuple[List[FolderJob], List[str]]:
    warnings = []
    subfolder_jobs = []

    def _onerror(err):
        warnings.append("Permission denied while scanning %s: %s" % (err.filename, err.strerror))

    for current_root, dirnames, filenames in os.walk(str(root), topdown=True, onerror=_onerror):
        dirnames.sort()
        filenames.sort()
        folder = Path(current_root)
        videos = [
            folder / filename
            for filename in filenames
            if (folder / filename).suffix.lower() in VIDEO_EXTS
        ]
        if videos:
            output_name = "preview_%s.docx" % safe_output_name(folder.name)
            subfolder_jobs.append(FolderJob(folder=folder, videos=videos, output_path=folder / output_name))

    subfolder_jobs.sort(key=lambda item: str(item.folder))
    if preview_scope == PreviewScope.SUBFOLDER:
        return subfolder_jobs, warnings

    project_videos = []
    project_warnings = []
    for job in subfolder_jobs:
        project_videos.extend(job.videos)
        project_warnings.extend(job.warnings)

    if not project_videos:
        return [], warnings

    project_output = root / ("preview_%s.docx" % safe_output_name(root.name))
    project_job = FolderJob(
        folder=root,
        videos=sorted(project_videos),
        output_path=project_output,
        warnings=project_warnings,
    )
    return [project_job], warnings
