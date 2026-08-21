import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import outlook_calendar_sync as syncmod

ROOT = Path(__file__).resolve().parents[1]
MACRO_FILE = ROOT / "OutlookSyncControls.bas"


class ProcessControlTests(unittest.TestCase):
    def test_status_uses_status_file_and_process_check(self):
        with tempfile.TemporaryDirectory() as tmp:
            status_path = Path(tmp) / "status.json"
            status_path.write_text(
                json.dumps(
                    {
                        "pid": 1234,
                        "started_at": "2026-01-02T03:04:05",
                        "state": "syncing",
                    }
                ),
                encoding="utf-8",
            )

            status = syncmod.get_running_status(
                status_path=status_path,
                is_running_func=lambda pid: pid == 1234,
            )

        self.assertEqual(status["pid"], 1234)
        self.assertEqual(status["state"], "syncing")

    def test_status_prints_waiting_state_and_last_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            status_path = Path(tmp) / "status.json"
            status_path.write_text(
                json.dumps(
                    {
                        "pid": 1234,
                        "started_at": "2026-01-02T03:04:05",
                        "state": "waiting",
                        "message": "Could not find Outlook store",
                    }
                ),
                encoding="utf-8",
            )
            stdout = io.StringIO()

            with contextlib.redirect_stdout(stdout):
                result = syncmod.print_status(
                    status_path=status_path,
                    is_running_func=lambda pid: pid == 1234,
                )

        self.assertEqual(result, 0)
        message = stdout.getvalue()
        self.assertIn("waiting", message)
        self.assertIn("Could not find Outlook store", message)

    def test_stop_requests_and_cleans_up_when_process_is_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            status_path = tmp_path / "status.json"
            stop_path = tmp_path / "stop.request"
            process_registry_path = tmp_path / "processes.json"
            status_path.write_text(json.dumps({"pid": 2222}), encoding="utf-8")
            alive = {2222}
            stdout = io.StringIO()

            with contextlib.redirect_stdout(stdout):
                result = syncmod.stop_sync(
                    status_path=status_path,
                    stop_path=stop_path,
                    process_registry_path=process_registry_path,
                    lock_path=tmp_path / "control.lock",
                    is_running_func=lambda pid: pid in alive,
                    terminate_func=lambda pid: alive.discard(pid) is None,
                    wait_seconds=0,
                )

            self.assertEqual(result, 0)
            self.assertTrue(stop_path.exists())
            self.assertFalse(status_path.exists())
            self.assertIn("Stopped", stdout.getvalue())

    def test_start_launches_background_run_when_not_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_path = tmp_path / "outlook_calendar_sync.py"
            script_path.write_text("", encoding="utf-8")
            settings_path = tmp_path / "real_settings.py"
            settings_path.write_text("", encoding="utf-8")
            status_path = tmp_path / "status.json"
            stop_path = tmp_path / "stop.request"
            control_log_path = tmp_path / "control.log"
            stdout = io.StringIO()

            with (
                mock.patch.object(syncmod.subprocess, "Popen") as popen,
                contextlib.redirect_stdout(stdout),
            ):
                popen.return_value.pid = 3333
                result = syncmod.start_sync(
                    script_path=script_path,
                    python_executable="python.exe",
                    settings_path=settings_path,
                    status_path=status_path,
                    stop_path=stop_path,
                    control_log_path=control_log_path,
                    process_registry_path=tmp_path / "processes.json",
                    lock_path=tmp_path / "control.lock",
                    is_running_func=lambda _pid: False,
                )

            self.assertEqual(result, 0)
            popen.assert_called_once()
            command = popen.call_args.args[0]
            self.assertEqual(command[-3:], ["--run", "--settings", str(settings_path)])
            self.assertIn(str(script_path), command)
            self.assertIn("Start requested", stdout.getvalue())

    def test_start_reports_already_running_from_registered_process(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_path = tmp_path / "outlook_calendar_sync.py"
            script_path.write_text("", encoding="utf-8")
            process_registry_path = tmp_path / "processes.json"
            process_registry_path.write_text(
                json.dumps({"pids": [4444]}),
                encoding="utf-8",
            )
            stdout = io.StringIO()

            with (
                mock.patch.object(syncmod.subprocess, "Popen") as popen,
                contextlib.redirect_stdout(stdout),
            ):
                result = syncmod.start_sync(
                    script_path=script_path,
                    process_registry_path=process_registry_path,
                    lock_path=tmp_path / "control.lock",
                    is_running_func=lambda pid: pid == 4444,
                )

            self.assertEqual(result, 0)
            popen.assert_not_called()
            self.assertIn("Already running", stdout.getvalue())

    def test_start_reports_busy_when_another_control_action_has_lock(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_path = tmp_path / "outlook_calendar_sync.py"
            script_path.write_text("", encoding="utf-8")
            lock_path = tmp_path / "control.lock"
            lock_path.write_text(json.dumps({"pid": 5555}), encoding="utf-8")
            stdout = io.StringIO()

            with (
                mock.patch.object(syncmod.subprocess, "Popen") as popen,
                contextlib.redirect_stdout(stdout),
            ):
                result = syncmod.start_sync(
                    script_path=script_path,
                    lock_path=lock_path,
                    lock_timeout_seconds=0,
                    is_running_func=lambda pid: pid == 5555,
                )

            self.assertEqual(result, 1)
            popen.assert_not_called()
            self.assertIn("already in progress", stdout.getvalue())

    def test_stop_terminates_registered_and_status_processes(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            status_path = tmp_path / "status.json"
            stop_path = tmp_path / "stop.request"
            process_registry_path = tmp_path / "processes.json"
            lock_path = tmp_path / "control.lock"
            status_path.write_text(json.dumps({"pid": 1111}), encoding="utf-8")
            process_registry_path.write_text(
                json.dumps({"pids": [2222, 3333]}),
                encoding="utf-8",
            )
            alive = {1111, 2222, 3333}
            terminated = []
            stdout = io.StringIO()

            def terminate(pid):
                terminated.append(pid)
                alive.discard(pid)
                return True

            with contextlib.redirect_stdout(stdout):
                result = syncmod.stop_sync(
                    status_path=status_path,
                    stop_path=stop_path,
                    process_registry_path=process_registry_path,
                    lock_path=lock_path,
                    is_running_func=lambda pid: pid in alive,
                    terminate_func=terminate,
                    wait_seconds=0,
                )

            self.assertEqual(result, 0)
            self.assertEqual(terminated, [1111, 2222, 3333])
            self.assertTrue(stop_path.exists())
            self.assertFalse(status_path.exists())
            self.assertEqual(json.loads(process_registry_path.read_text())["pids"], [])
            self.assertIn("Stopped OutlookCalendarSync", stdout.getvalue())

    def test_restart_refuses_when_sync_is_not_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_path = tmp_path / "outlook_calendar_sync.py"
            script_path.write_text("", encoding="utf-8")
            stdout = io.StringIO()

            with (
                mock.patch.object(syncmod.subprocess, "Popen") as popen,
                contextlib.redirect_stdout(stdout),
            ):
                result = syncmod.restart_sync(
                    script_path=script_path,
                    status_path=tmp_path / "status.json",
                    stop_path=tmp_path / "stop.request",
                    process_registry_path=tmp_path / "processes.json",
                    lock_path=tmp_path / "control.lock",
                    is_running_func=lambda _pid: False,
                )

            self.assertEqual(result, 1)
            popen.assert_not_called()
            self.assertIn("not running", stdout.getvalue())
            self.assertIn("Start", stdout.getvalue())

    def test_restart_stops_then_starts_when_sync_is_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            script_path = tmp_path / "outlook_calendar_sync.py"
            script_path.write_text("", encoding="utf-8")
            status_path = tmp_path / "status.json"
            status_path.write_text(json.dumps({"pid": 6666}), encoding="utf-8")
            process_registry_path = tmp_path / "processes.json"
            process_registry_path.write_text(
                json.dumps({"pids": [6666]}),
                encoding="utf-8",
            )
            alive = {6666}
            stdout = io.StringIO()

            def terminate(pid):
                alive.discard(pid)
                return True

            with (
                mock.patch.object(syncmod.subprocess, "Popen") as popen,
                contextlib.redirect_stdout(stdout),
            ):
                popen.return_value.pid = 7777
                result = syncmod.restart_sync(
                    script_path=script_path,
                    python_executable="python.exe",
                    status_path=status_path,
                    stop_path=tmp_path / "stop.request",
                    process_registry_path=process_registry_path,
                    lock_path=tmp_path / "control.lock",
                    control_log_path=tmp_path / "control.log",
                    is_running_func=lambda pid: pid in alive,
                    terminate_func=terminate,
                    wait_seconds=0,
                )

            self.assertEqual(result, 0)
            popen.assert_called_once()
            self.assertIn("Restarted OutlookCalendarSync", stdout.getvalue())

    def test_run_exits_when_stop_request_exists(self):
        class FakeSync:
            settings = SimpleNamespace(POLL_SECONDS=5)

            def __init__(self):
                self.reconcile_calls = 0

            def reconcile(self):
                self.reconcile_calls += 1

        fake = FakeSync()

        with tempfile.TemporaryDirectory() as tmp:
            stop_path = Path(tmp) / "stop.request"
            stop_path.write_text("stop", encoding="utf-8")
            stdout = io.StringIO()

            with (
                mock.patch.object(syncmod, "STOP_REQUEST_PATH", stop_path),
                contextlib.redirect_stdout(stdout),
            ):
                syncmod.OutlookSync.run(fake)

        self.assertEqual(fake.reconcile_calls, 0)
        self.assertIn("stopped", stdout.getvalue().lower())

    def test_wait_for_outlook_ready_records_waiting_state(self):
        with tempfile.TemporaryDirectory() as tmp:
            status_path = Path(tmp) / "status.json"
            attempts = []

            def factory(settings, com_modules):
                attempts.append(1)
                if len(attempts) == 1:
                    raise RuntimeError("stores are still loading")
                return object()

            clock = SimpleNamespace(now=0)

            def monotonic():
                return clock.now

            def sleep(seconds):
                clock.now += seconds

            result = syncmod.wait_for_outlook_ready(
                settings=SimpleNamespace(
                    OUTLOOK_READY_TIMEOUT_SECONDS=30,
                    OUTLOOK_READY_RETRY_SECONDS=5,
                ),
                com_modules=(object(), object()),
                settings_path=Path("C:/calendar/settings.py"),
                status_path=status_path,
                sync_factory=factory,
                sleep_func=sleep,
                monotonic_func=monotonic,
            )

            status = json.loads(status_path.read_text(encoding="utf-8"))

        self.assertIsNotNone(result)
        self.assertEqual(status["state"], "waiting")
        self.assertIn("stores are still loading", status["message"])

    def test_wait_for_outlook_ready_honors_stop_request_before_outlook_ready(self):
        with tempfile.TemporaryDirectory() as tmp:
            stop_path = Path(tmp) / "stop.request"
            stop_path.write_text("stop", encoding="utf-8")

            def factory(settings, com_modules):
                raise AssertionError("Outlook startup should not be attempted after stop")

            with self.assertRaises(KeyboardInterrupt):
                syncmod.wait_for_outlook_ready(
                    settings=SimpleNamespace(
                        OUTLOOK_READY_TIMEOUT_SECONDS=30,
                        OUTLOOK_READY_RETRY_SECONDS=5,
                    ),
                    com_modules=(object(), object()),
                    stop_path=stop_path,
                    sync_factory=factory,
                )

    def test_main_status_does_not_require_pywin32(self):
        stderr = io.StringIO()
        stdout = io.StringIO()

        with (
            mock.patch.object(sys, "argv", ["outlook_calendar_sync.py", "--status"]),
            mock.patch.object(syncmod, "print_status", return_value=0) as print_status,
            mock.patch.object(syncmod, "load_outlook_com_modules") as load_com,
            contextlib.redirect_stdout(stdout),
            contextlib.redirect_stderr(stderr),
        ):
            result = syncmod.main()

        self.assertEqual(result, 0)
        print_status.assert_called_once()
        load_com.assert_not_called()
        self.assertEqual(stderr.getvalue(), "")

    def test_main_start_passes_settings_path_without_pywin32(self):
        settings_path = Path("C:/calendar/settings.py")

        with (
            mock.patch.object(
                sys,
                "argv",
                ["outlook_calendar_sync.py", "--start", "--settings", str(settings_path)],
            ),
            mock.patch.object(syncmod, "start_sync", return_value=0) as start_sync,
            mock.patch.object(syncmod, "load_outlook_com_modules") as load_com,
        ):
            result = syncmod.main()

        self.assertEqual(result, 0)
        start_sync.assert_called_once()
        self.assertEqual(start_sync.call_args.kwargs["settings_path"], settings_path)
        load_com.assert_not_called()

    def test_main_foreground_run_loads_explicit_settings_path(self):
        settings_path = Path("C:/calendar/settings.py")

        class FakeSync:
            calendar_a = SimpleNamespace(FolderPath="Calendar A")
            calendar_b = SimpleNamespace(FolderPath="Calendar B")

            def run(self):
                return None

        with tempfile.TemporaryDirectory() as tmp:
            status_path = Path(tmp) / "status.json"

            with (
                mock.patch.object(
                    sys,
                    "argv",
                    ["outlook_calendar_sync.py", "--run", "--settings", str(settings_path)],
                ),
                mock.patch.object(syncmod, "STATUS_PATH", status_path),
                mock.patch.object(syncmod, "setup_logging"),
                mock.patch.object(syncmod, "mark_running", return_value={"pid": 1}),
                mock.patch.object(syncmod, "clear_running_status"),
                mock.patch.object(syncmod, "load_settings", return_value=object()) as load_settings,
                mock.patch.object(
                    syncmod,
                    "load_outlook_com_modules",
                    return_value=(object(), object()),
                ),
                mock.patch.object(
                    syncmod,
                    "wait_for_outlook_ready",
                    return_value=FakeSync(),
                ),
            ):
                result = syncmod.main()

            self.assertEqual(result, 0)
            load_settings.assert_called_once_with(settings_path)
            status = json.loads(status_path.read_text(encoding="utf-8"))
            self.assertEqual(status["state"], "syncing")
            self.assertIn("Calendar A", status["message"])

    def test_main_clears_status_file_when_startup_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            status_path = tmp_path / "status.json"
            stop_path = tmp_path / "stop.request"
            stderr = io.StringIO()

            with (
                mock.patch.object(sys, "argv", ["outlook_calendar_sync.py"]),
                mock.patch.object(syncmod, "STATUS_PATH", status_path),
                mock.patch.object(syncmod, "STOP_REQUEST_PATH", stop_path),
                mock.patch.object(syncmod, "PROCESS_REGISTRY_PATH", tmp_path / "processes.json"),
                mock.patch.object(syncmod, "setup_logging"),
                mock.patch.object(syncmod, "load_settings", return_value=object()),
                mock.patch.object(
                    syncmod,
                    "load_outlook_com_modules",
                    return_value=(object(), object()),
                ),
                mock.patch.object(
                    syncmod,
                    "wait_for_outlook_ready",
                    side_effect=RuntimeError("startup failed"),
                ),
                contextlib.redirect_stderr(stderr),
            ):
                result = syncmod.main()

            self.assertEqual(result, 1)
            self.assertIn("startup failed", stderr.getvalue())
            self.assertFalse(status_path.exists())

    def test_macro_module_exposes_start_stop_status_restart_buttons(self):
        text = MACRO_FILE.read_text(encoding="utf-8")

        self.assertIn("Public Sub OutlookSync_Start()", text)
        self.assertIn("Public Sub OutlookSync_Stop()", text)
        self.assertIn("Public Sub OutlookSync_Status()", text)
        self.assertIn("Public Sub OutlookSync_Restart()", text)
        self.assertIn('MsgBox RunAndCapture("--start")', text)
        self.assertIn('MsgBox RunAndCapture("--stop")', text)
        self.assertIn('MsgBox RunAndCapture("--restart")', text)
        self.assertIn("SETTINGS_FILE", text)
        self.assertIn("--start", text)
        self.assertIn("--stop", text)
        self.assertIn("--status", text)
        self.assertIn("--restart", text)
        self.assertIn("--settings", text)

    def test_macro_captures_command_output_without_visible_console(self):
        text = MACRO_FILE.read_text(encoding="utf-8")

        self.assertNotIn("shell.Exec", text)
        self.assertIn("shell.Run command, 0, True", text)
        self.assertIn(" 2>&1", text)
        self.assertIn("GetTempName", text)
        self.assertIn("DeleteFile outputPath, True", text)

    def test_macro_uses_real_python_executable_instead_of_python_launcher(self):
        text = MACRO_FILE.read_text(encoding="utf-8")

        self.assertIn(
            'Private Const PYTHON_COMMAND As String = "%LOCALAPPDATA%\\Programs\\Python\\Python312\\python.exe"',
            text,
        )
        self.assertIn("Quote(ExpandEnv(PYTHON_COMMAND))", text)


if __name__ == "__main__":
    unittest.main()
