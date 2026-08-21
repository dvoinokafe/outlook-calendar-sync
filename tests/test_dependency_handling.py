import os
import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import outlook_calendar_sync as syncmod

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "outlook_calendar_sync.py"


def env_with_user_dirs_under(tmp_path, localappdata=None):
    env = os.environ.copy()
    env["APPDATA"] = str(tmp_path / "appdata")
    env["LOCALAPPDATA"] = str(localappdata or tmp_path / "localappdata")
    env["PYTHONNOUSERSITE"] = "1"
    env.pop("PYTHONPATH", None)
    return env


class DependencyHandlingTests(unittest.TestCase):
    def run_without_site_packages(self, tmp_path, *args, localappdata=None):
        return subprocess.run(
            [sys.executable, "-S", str(SCRIPT), *args],
            cwd=ROOT,
            env=env_with_user_dirs_under(tmp_path, localappdata=localappdata),
            capture_output=True,
            text=True,
            timeout=10,
        )

    def test_missing_pywin32_exits_with_install_guidance(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = self.run_without_site_packages(Path(tmp))

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("pywin32", result.stderr)
        self.assertIn("pythoncom", result.stderr)
        self.assertIn("win32com.client", result.stderr)
        self.assertIn(
            f'"{sys.executable}" -m pip install --user -r requirements.txt',
            result.stderr,
        )
        self.assertNotIn("Traceback", result.stderr)

    def test_missing_pywin32_guidance_precedes_log_directory_setup(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            blocked_localappdata = tmp_path / "localappdata-is-a-file"
            blocked_localappdata.write_text("", encoding="utf-8")

            result = self.run_without_site_packages(
                tmp_path,
                localappdata=blocked_localappdata,
            )

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("pywin32", result.stderr)
        self.assertIn("pythoncom", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_install_autostart_does_not_require_pywin32(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            result = self.run_without_site_packages(tmp_path, "--install-autostart")

            startup_cmd = (
                tmp_path
                / "appdata"
                / "Microsoft"
                / "Windows"
                / "Start Menu"
                / "Programs"
                / "Startup"
                / "OutlookCalendarSync.cmd"
            )

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(startup_cmd.exists())
            launcher = startup_cmd.read_text(encoding="utf-8")
            self.assertIn(str(SCRIPT), launcher)
            self.assertIn("--start", launcher)
            self.assertNotIn("Traceback", result.stderr)

    def test_logging_setup_error_prints_clean_message(self):
        stderr = io.StringIO()

        with (
            mock.patch.object(sys, "argv", [str(SCRIPT)]),
            mock.patch.object(syncmod, "load_settings", return_value=object()),
            mock.patch.object(syncmod, "load_outlook_com_modules", return_value=(object(), object())),
            mock.patch.object(
                syncmod,
                "setup_logging",
                side_effect=PermissionError("[WinError 5] Access is denied"),
            ),
            contextlib.redirect_stderr(stderr),
        ):
            result = syncmod.main()

        message = stderr.getvalue()
        self.assertEqual(result, 1)
        self.assertIn("OutlookCalendarSync stopped because of an error.", message)
        self.assertIn("PermissionError", message)
        self.assertIn("Access is denied", message)
        self.assertNotIn("Traceback", message)

    def test_reconciliation_error_prints_clean_message(self):
        class FailingSync:
            settings = SimpleNamespace(POLL_SECONDS=1)

            def reconcile(self):
                raise RuntimeError("copy failed")

        stderr = io.StringIO()

        with tempfile.TemporaryDirectory() as tmp:
            with (
                mock.patch.object(syncmod, "SYNC_LOCK_PATH", Path(tmp) / "sync.lock"),
                mock.patch.object(syncmod.time, "sleep", side_effect=KeyboardInterrupt),
                contextlib.redirect_stderr(stderr),
                self.assertRaises(KeyboardInterrupt),
            ):
                syncmod.OutlookSync.run(FailingSync())

        message = stderr.getvalue()
        self.assertIn("OutlookCalendarSync stopped because of an error.", message)
        self.assertIn("RuntimeError", message)
        self.assertIn("copy failed", message)
        self.assertNotIn("Traceback", message)


if __name__ == "__main__":
    unittest.main()
