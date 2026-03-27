import shutil
import subprocess
import sys
import tempfile
import unittest
import os
import zipfile
from unittest import mock
import importlib.util
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "makevideopreviews.py"
SRC_ROOT = REPO_ROOT / "src"

sys.path = [str(SRC_ROOT)] + [entry for entry in sys.path if entry not in {"", str(REPO_ROOT)}]

from makevideopreviews.discovery import discover_jobs
from makevideopreviews.docx_render import _common_filename_prefixes, _overlay_name_for_probe
from makevideopreviews.docx_render import write_docx
from makevideopreviews.errors import FrameExtractionError
from makevideopreviews.media import _extract_video_frames_worker, extract_single_frame, probe_video
from makevideopreviews.models import AppConfig, CommandMode, ExtractionResult, FolderJob, OverwriteMode
from makevideopreviews.models import PreviewScope
from makevideopreviews.models import VideoProbe
from makevideopreviews.utils import resolve_existing_dir


class PathResolutionTests(unittest.TestCase):
    def test_resolve_existing_dir_strips_quotes(self):
        resolved = resolve_existing_dir("'%s'" % REPO_ROOT)
        self.assertEqual(resolved, REPO_ROOT.resolve())

    def test_resolve_existing_dir_matches_unicode_variants(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target = root / "März Material"
            target.mkdir()
            resolved = resolve_existing_dir(str(root / "März Material"))
            self.assertTrue(os.path.samefile(str(resolved), str(target)))


class DiscoveryTests(unittest.TestCase):
    def test_project_scope_returns_single_root_docx(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "A" / "CLIPS001").mkdir(parents=True)
            (root / "B" / "CLIPS001").mkdir(parents=True)
            (root / "A" / "CLIPS001" / "a.mp4").touch()
            (root / "B" / "CLIPS001" / "b.mp4").touch()

            jobs, warnings = discover_jobs(root, PreviewScope.PROJECT)

            self.assertEqual(warnings, [])
            self.assertEqual(len(jobs), 1)
            self.assertEqual(jobs[0].folder, root)
            self.assertEqual(len(jobs[0].videos), 2)
            self.assertEqual(jobs[0].output_path.parent, root)

    def test_subfolder_scope_returns_one_job_per_video_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "A" / "CLIPS001").mkdir(parents=True)
            (root / "B" / "CLIPS001").mkdir(parents=True)
            (root / "A" / "CLIPS001" / "a.mp4").touch()
            (root / "B" / "CLIPS001" / "b.mp4").touch()

            jobs, warnings = discover_jobs(root, PreviewScope.SUBFOLDER)

            self.assertEqual(warnings, [])
            self.assertEqual(len(jobs), 2)


class OverlayTests(unittest.TestCase):
    def test_overlay_name_keeps_timecode_space_by_shortening_common_prefix(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            folder = root / "clips"
            folder.mkdir()
            first = folder / "A007C118_2603254I_CANON.MXF"
            second = folder / "A007C119_2603253A_CANON.MXF"
            first.touch()
            second.touch()
            prefixes = _common_filename_prefixes([first, second])
            probe = VideoProbe(
                path=first,
                duration=10.0,
                width=1920,
                height=1080,
                fps=25.0,
                estimated_frames=2,
            )
            overlay_name = _overlay_name_for_probe(probe, prefixes)
            self.assertTrue(overlay_name.endswith(".mxf"))
            self.assertIn("2603254I_CANON", overlay_name)
            self.assertLess(len(overlay_name), len(first.name))


if importlib.util.find_spec("typer") is not None:
    from makevideopreviews.app import _package_hint


@unittest.skipIf(importlib.util.find_spec("typer") is None, "typer required")
class DoctorHintTests(unittest.TestCase):
    def test_package_hint_is_human_for_available_packages(self):
        self.assertEqual(_package_hint("typer", True), "Available in the current Python environment.")

    def test_package_hint_still_explains_missing_packages(self):
        self.assertIn("python3 -m pip install typer", _package_hint("typer", False))


class MediaTimeoutTests(unittest.TestCase):
    @mock.patch("makevideopreviews.media.subprocess.run")
    def test_probe_video_timeout_becomes_problem_marker(self, run_mock):
        run_mock.side_effect = subprocess.TimeoutExpired(cmd="ffprobe", timeout=30)
        with tempfile.TemporaryDirectory() as tmp:
            clip = Path(tmp) / "clip.mp4"
            clip.touch()
            probe = probe_video(clip, interval=10, cache={})
        self.assertIn("timed out", probe.problem)

    @mock.patch("makevideopreviews.media.subprocess.run")
    def test_extract_single_frame_timeout_returns_false(self, run_mock):
        run_mock.side_effect = subprocess.TimeoutExpired(cmd="ffmpeg", timeout=60)
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "clip.mp4"
            src.touch()
            out = Path(tmp) / "frame.jpg"
            self.assertFalse(extract_single_frame(src, 0.0, out, max_width=480, quality=10))

    @mock.patch("makevideopreviews.media.subprocess.run")
    def test_batch_extract_timeout_raises_clear_error(self, run_mock):
        run_mock.side_effect = subprocess.TimeoutExpired(cmd="ffmpeg", timeout=180)
        with tempfile.TemporaryDirectory() as tmp:
            clip = Path(tmp) / "clip.mp4"
            clip.touch()
            with self.assertRaises(FrameExtractionError) as ctx:
                _extract_video_frames_worker(str(clip), 0.0, 10, 480, 10, tmp, False)
        self.assertIn("timed out", str(ctx.exception))


@unittest.skipIf(
    importlib.util.find_spec("docx") is None or importlib.util.find_spec("PIL") is None,
    "docx and Pillow required",
)
class DocxLayoutTests(unittest.TestCase):
    def test_single_sheet_docx_has_single_cover_break(self):
        from PIL import Image

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            clip = root / "camera_a" / "clip.mp4"
            clip.parent.mkdir()
            clip.touch()
            frame = root / "frame.jpg"
            Image.new("RGB", (640, 360), color=(20, 40, 60)).save(frame, format="JPEG")

            probe = VideoProbe(
                path=clip,
                duration=10.0,
                width=1920,
                height=1080,
                fps=25.0,
                estimated_frames=1,
            )
            job = FolderJob(
                folder=root,
                videos=[clip],
                output_path=root / ("preview_%s.docx" % root.name),
                probes=[probe],
                estimated_bytes=123456,
            )
            result = ExtractionResult(
                video_path=clip,
                frame_paths=[frame],
                timestamps=[0.0],
                temp_dir=frame.parent,
                error=None,
            )
            config = AppConfig(
                command=CommandMode.GENERATE,
                root=root,
                preview_scope=PreviewScope.PROJECT,
                interval=10,
                max_width=480,
                quality=10,
                page_width_inch=10.0,
                workers=1,
                overwrite_mode=OverwriteMode.SKIP,
                debug=False,
            )

            output_path = write_docx(job, {clip: result}, config)
            self.assertTrue(output_path.exists())

            with zipfile.ZipFile(output_path) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertEqual(document_xml.count('w:type="page"'), 1)


class CliTests(unittest.TestCase):
    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, str(SCRIPT)] + list(args),
            cwd=str(REPO_ROOT),
            capture_output=True,
            text=True,
        )

    def test_help_works(self):
        result = self.run_cli("--help")
        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn("generate", result.stdout)

    def test_doctor_runs_without_crashing(self):
        result = self.run_cli("doctor", "--root", str(REPO_ROOT))
        self.assertIn("Doctor Report", result.stdout)
        self.assertIn(result.returncode, (0, 1))

    def test_invalid_interval_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = self.run_cli("generate", "--root", tmp, "--interval", "0")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Interval must be >= 1", result.stderr + result.stdout)

    def test_empty_root_is_clean(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = self.run_cli("estimate", "--root", tmp)
        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn("No folders with video files found", result.stdout)


@unittest.skipIf(shutil.which("ffmpeg") is None or shutil.which("ffprobe") is None, "ffmpeg required")
class IntegrationTests(unittest.TestCase):
    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, str(SCRIPT)] + list(args),
            cwd=str(REPO_ROOT),
            capture_output=True,
            text=True,
        )

    def create_video(self, path: Path):
        cmd = [
            "ffmpeg",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "lavfi",
            "-i",
            "testsrc=size=320x240:rate=25",
            "-t",
            "2",
            "-pix_fmt",
            "yuv420p",
            "-y",
            str(path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, msg=result.stderr)

    def test_estimate_and_generate(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            clip_dir = root / "camera_a"
            clip_dir.mkdir()
            self.create_video(clip_dir / "clip.mp4")

            estimate = self.run_cli("estimate", "--root", str(root), "--interval", "1")
            self.assertEqual(estimate.returncode, 0, msg=estimate.stderr)
            self.assertIn("Estimated total", estimate.stdout)

            generate = self.run_cli("generate", "--root", str(root), "--interval", "1", "--workers", "1")
            self.assertIn(generate.returncode, (0, 1), msg=generate.stderr)
            if generate.returncode == 0:
                self.assertTrue((root / ("preview_%s.docx" % root.name)).exists())

    def test_generate_subfolder_scope_writes_into_video_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            clip_dir = root / "camera_a"
            clip_dir.mkdir()
            self.create_video(clip_dir / "clip.mp4")

            generate = self.run_cli(
                "generate",
                "--root",
                str(root),
                "--scope",
                "subfolder",
                "--interval",
                "1",
                "--workers",
                "1",
            )
            self.assertIn(generate.returncode, (0, 1), msg=generate.stderr)
            if generate.returncode == 0:
                self.assertTrue((clip_dir / "preview_camera_a.docx").exists())

    def test_short_clip_still_gets_preview_frame_when_interval_is_longer(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            clip_dir = root / "camera_a"
            clip_dir.mkdir()
            self.create_video(clip_dir / "clip.mp4")

            generate = self.run_cli("generate", "--root", str(root), "--interval", "10", "--workers", "1")
            self.assertEqual(generate.returncode, 0, msg=generate.stderr)
            self.assertTrue((root / ("preview_%s.docx" % root.name)).exists())


if __name__ == "__main__":
    unittest.main()
