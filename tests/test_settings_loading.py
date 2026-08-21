import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "outlook_calendar_sync.py"


SETTINGS_TEXT = """
STORE_A_HINT = "alpha@example.com"
STORE_B_HINT = "bravo@example.com"
DAYS_BACK = 7
DAYS_FORWARD = 45
POLL_SECONDS = 180
SCAN_BATCH_SIZE = 12
SCAN_BATCH_PAUSE_SECONDS = 0.2
OUTLOOK_READY_TIMEOUT_SECONDS = 42
OUTLOOK_READY_RETRY_SECONDS = 3
"""


def copy_script_to(tmp_path):
    script_copy = tmp_path / "outlook_calendar_sync.py"
    shutil.copy2(SCRIPT, script_copy)
    return script_copy


def import_script_copy(script_copy):
    spec = importlib.util.spec_from_file_location(
        "outlook_calendar_sync_settings_test",
        script_copy,
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def isolated_env(tmp_path):
    env = os.environ.copy()
    env["APPDATA"] = str(tmp_path / "appdata")
    env["LOCALAPPDATA"] = str(tmp_path / "localappdata")
    env["PYTHONNOUSERSITE"] = "1"
    env.pop("PYTHONPATH", None)
    return env


class SettingsLoadingTests(unittest.TestCase):
    def test_settings_file_controls_accounts_and_periods(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_copy = copy_script_to(tmp_path)
            (tmp_path / "settings.py").write_text(SETTINGS_TEXT, encoding="utf-8")
            module = import_script_copy(script_copy)

            loader = getattr(module, "load_settings", None)
            self.assertIsNotNone(loader)
            settings = loader()

        self.assertEqual(settings.STORE_A_HINT, "alpha@example.com")
        self.assertEqual(settings.STORE_B_HINT, "bravo@example.com")
        self.assertEqual(settings.DAYS_BACK, 7)
        self.assertEqual(settings.DAYS_FORWARD, 45)
        self.assertEqual(settings.POLL_SECONDS, 180)
        self.assertEqual(settings.SCAN_BATCH_SIZE, 12)
        self.assertEqual(settings.SCAN_BATCH_PAUSE_SECONDS, 0.2)
        self.assertEqual(settings.OUTLOOK_READY_TIMEOUT_SECONDS, 42)
        self.assertEqual(settings.OUTLOOK_READY_RETRY_SECONDS, 3)

    def test_missing_settings_file_exits_with_clear_guidance_before_pywin32(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_copy = copy_script_to(tmp_path)
            result = subprocess.run(
                [sys.executable, "-S", str(script_copy)],
                cwd=tmp_path,
                env=isolated_env(tmp_path),
                capture_output=True,
                text=True,
                timeout=10,
            )

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("settings.py", result.stderr)
        self.assertIn("STORE_A_HINT", result.stderr)
        self.assertNotIn("pywin32", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_install_autostart_does_not_require_settings_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_copy = copy_script_to(tmp_path)
            result = subprocess.run(
                [sys.executable, "-S", str(script_copy), "--install-autostart"],
                cwd=tmp_path,
                env=isolated_env(tmp_path),
                capture_output=True,
                text=True,
                timeout=10,
            )

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
            self.assertIn("--start", startup_cmd.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
