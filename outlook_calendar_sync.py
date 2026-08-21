r"""
outlook_calendar_sync.py

NO-ADMIN, PER-USER CALENDAR SYNCHRONIZER FOR CLASSIC OUTLOOK
============================================================

WHAT THIS IS
------------
This script synchronizes two calendars that are already available in the
same Classic Outlook for Windows profile.

It is intended for environments where:

- you have two company Microsoft 365 accounts,
- they belong to different tenants,
- tenant admin consent for Microsoft Graph will NOT be granted,
- local administrator rights will NOT be granted,
- both accounts are already configured in Classic Outlook.

It does NOT install a Windows Service.
Instead, it runs in your interactive Windows user session and can be
started automatically at logon without elevation.

HOW IT WORKS
------------
The script automates Classic Outlook through the Outlook COM object model.

It keeps a local SQLite database mapping mirrored calendar occurrences
between the two company calendars.

It synchronizes:

- creation
- change
- deletion
- recurring meetings
- one changed occurrence in a recurring series
- one deleted occurrence in a recurring series

Recurring meetings are synchronized occurrence-by-occurrence inside a
rolling window. This avoids having to reproduce recurrence exception
internals across two unrelated company mailboxes.

IMPORTANT
---------
This script requires CLASSIC OUTLOOK.

It does not work with New Outlook.

Outlook must be running, or able to start, in the signed-in desktop
session.

This script should NOT be run as a Windows Service.


DEPENDENCIES
============
Install for the current user:

    py -m pip install --user -r requirements.txt

The pywin32 package provides both pythoncom and win32com.client. Do not
install pythoncom separately.

Python includes sqlite3, json, hashlib, pathlib, and datetime.


CONFIGURATION
=============
Edit the values in settings.py, which must sit next to this script:

    STORE_A_HINT
    STORE_B_HINT

Set them to distinctive text that appears in the corresponding Outlook
store/mailbox display names, usually the email address.

Example:

    STORE_A_HINT = "me@company-a.com"
    STORE_B_HINT = "me@company-b.com"

The script searches Outlook stores case-insensitively.

Optional settings:

    DAYS_BACK = 2
    DAYS_FORWARD = 2
    POLL_SECONDS = 5
    SCAN_BATCH_SIZE = 25
    SCAN_BATCH_PAUSE_SECONDS = 0.05
    OUTLOOK_READY_TIMEOUT_SECONDS = 300
    OUTLOOK_READY_RETRY_SECONDS = 10

The rolling window is necessary because Outlook recurrence expansion is
date-window based. The polling and scan-pause defaults favor Outlook UI
responsiveness over near-instant synchronization.


AUTOSTART WITHOUT ADMIN RIGHTS
==============================
Run:

    py outlook_calendar_sync.py --install-autostart

This creates a .cmd launcher in your per-user Windows Startup folder.
At Windows login, the script waits up to OUTLOOK_READY_TIMEOUT_SECONDS
for Classic Outlook and both configured calendars to become ready.

Remove it with:

    py outlook_calendar_sync.py --remove-autostart

No administrator rights are required.


MANUAL CONTROL
==============
The synchronizer can be controlled without admin rights:

    py outlook_calendar_sync.py --start
    py outlook_calendar_sync.py --stop
    py outlook_calendar_sync.py --status
    py outlook_calendar_sync.py --restart
    py outlook_calendar_sync.py --sync-selected

Use --settings "C:\Path\To\settings.py" when the real settings file is
not next to this script.

Start prints "Already running" and exits when a tracked sync process is
alive. Stop requests a clean shutdown, then stops any tracked sync
process that remains alive. Restart only works when sync is already
running; otherwise it asks you to start sync first.

Sync selected attaches to the currently open or selected appointment and
copies or updates only that item. It works even when the background
synchronizer is not running.

These commands are serialized by a per-user control lock and are
intended for command-line use and for Classic Outlook VBA buttons. See
OutlookSyncControls.bas.


TEST FIRST
==========
Before enabling autostart, run:

    py outlook_calendar_sync.py

Then test:

1. Create a normal appointment in Calendar A.
2. Confirm a copy appears in Calendar B.
3. Change it.
4. Confirm the copy changes.
5. Delete it.
6. Confirm the copy is deleted.
7. Repeat in the opposite direction.
8. Create a recurring series.
9. Change only one occurrence.
10. Confirm only that occurrence changes in the other calendar.
11. Delete only one occurrence.
12. Confirm only that mirrored occurrence is deleted.


DATA AND LOG FILES
==================
Per-user files are stored under:

    %LOCALAPPDATA%\OutlookCalendarSync\

Files:

    sync.db
    sync.log
    status.json
    stop.request
    processes.json
    control.log
    sync.lock

Deleting sync.db destroys the mapping between paired appointments and may
cause duplicates.


WHAT IS COPIED
==============
Copied fields:

- Subject
- Body
- Start
- End
- All-day status
- Location
- Categories
- Sensitivity
- Busy status
- Importance
- Reminder settings

NOT copied:

- organizer
- attendees / recipients
- meeting response state
- online-meeting metadata

This prevents the mirrored copy from becoming another organizer and
sending invitations or updates.

The mirrored item is a normal appointment copy.


CONFLICT POLICY
===============
The script keeps hashes for both sides.

If only one side changes, that side wins.

If both sides change before the next reconciliation pass, the item with
the newer LastModificationTime wins.


RECURRENCE MODEL
================
The script expands recurring items with:

    Items.IncludeRecurrences = True

inside the configured date window.

Each occurrence is assigned a stable logical key based on:

- series/global identity when available
- original occurrence start time

This lets an exception occurrence be treated independently from the rest
of the series.

A changed occurrence maps to only its paired occurrence.

A deleted occurrence disappears from the expanded calendar view and its
paired occurrence is then deleted during reconciliation.


KNOWN LIMITATIONS
=================
1. Outlook COM automation is interactive-user automation, not server-side.
2. Outlook may display security prompts depending on corporate policy.
3. Outlook item events are not perfectly reliable for bulk changes, so this
   script uses periodic reconciliation as the source of truth.
4. Synchronization only covers the configured rolling window.
5. Very large calendars can make reconciliation slower.
6. Moving occurrences outside the rolling window may look like deletion
   until they enter the window again.
7. Meeting organizer/attendee semantics are deliberately not replicated.


"""

import argparse
import hashlib
import importlib
import importlib.util
import json
import logging
import os
import sqlite3
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

# ---------- CONSTANTS ----------

OL_FOLDER_CALENDAR = 9
OL_APPOINTMENT_ITEM = 1
OL_RECURRENCE_STATE_NOT_RECURRING = 0

BASE_DIR = Path(os.environ.get("LOCALAPPDATA", ".")) / "OutlookCalendarSync"
DB_PATH = BASE_DIR / "sync.db"
LOG_PATH = BASE_DIR / "sync.log"
SETTINGS_PATH = Path(__file__).with_name("settings.py")

STARTUP_DIR = Path(os.environ.get("APPDATA", ".")) / (
    r"Microsoft\Windows\Start Menu\Programs\Startup"
)
STARTUP_CMD = STARTUP_DIR / "OutlookCalendarSync.cmd"
STATUS_PATH = BASE_DIR / "status.json"
STOP_REQUEST_PATH = BASE_DIR / "stop.request"
CONTROL_LOG_PATH = BASE_DIR / "control.log"
PROCESS_REGISTRY_PATH = BASE_DIR / "processes.json"
CONTROL_LOCK_PATH = BASE_DIR / "control.lock"
SYNC_LOCK_PATH = BASE_DIR / "sync.lock"
CONTROL_SLEEP_SECONDS = 0.25
CONTROL_LOCK_TIMEOUT_SECONDS = 10
SYNC_LOCK_TIMEOUT_SECONDS = 10
STOP_WAIT_SECONDS = 5

SYNC_PROP = "OutlookCalendarSyncKey"
SYNC_NS = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/"
LOGGING_READY = False

REQUIRED_SETTINGS = (
    "STORE_A_HINT",
    "STORE_B_HINT",
    "DAYS_BACK",
    "DAYS_FORWARD",
    "POLL_SECONDS",
    "SCAN_BATCH_SIZE",
    "SCAN_BATCH_PAUSE_SECONDS",
    "OUTLOOK_READY_TIMEOUT_SECONDS",
    "OUTLOOK_READY_RETRY_SECONDS",
)

SETTINGS_TEMPLATE = """\
STORE_A_HINT = "me@company-a.com"
STORE_B_HINT = "me@company-b.com"

DAYS_BACK = 2
DAYS_FORWARD = 2
POLL_SECONDS = 5
SCAN_BATCH_SIZE = 25
SCAN_BATCH_PAUSE_SECONDS = 0.05
OUTLOOK_READY_TIMEOUT_SECONDS = 300
OUTLOOK_READY_RETRY_SECONDS = 10
"""


class SettingsError(RuntimeError):
    pass


class SyncSettings:
    def __init__(
        self,
        store_a_hint,
        store_b_hint,
        days_back,
        days_forward,
        poll_seconds,
        scan_batch_size,
        scan_batch_pause_seconds,
        outlook_ready_timeout_seconds,
        outlook_ready_retry_seconds,
    ):
        self.STORE_A_HINT = store_a_hint
        self.STORE_B_HINT = store_b_hint
        self.DAYS_BACK = days_back
        self.DAYS_FORWARD = days_forward
        self.POLL_SECONDS = poll_seconds
        self.SCAN_BATCH_SIZE = scan_batch_size
        self.SCAN_BATCH_PAUSE_SECONDS = scan_batch_pause_seconds
        self.OUTLOOK_READY_TIMEOUT_SECONDS = outlook_ready_timeout_seconds
        self.OUTLOOK_READY_RETRY_SECONDS = outlook_ready_retry_seconds


def format_settings_error(message):
    return (
        "Missing or invalid settings.py.\n\n"
        f"Create a settings.py file next to {Path(__file__).name} with these values:\n\n"
        f"{SETTINGS_TEMPLATE}\n"
        f"{message}"
    )


def normalize_text_setting(module, name):
    value = getattr(module, name)
    if not isinstance(value, str) or not value.strip():
        raise SettingsError(format_settings_error(f"{name} must be a non-empty string."))
    return value.strip()


def normalize_int_setting(module, name, minimum):
    value = getattr(module, name)
    if isinstance(value, bool):
        raise SettingsError(format_settings_error(f"{name} must be an integer."))
    try:
        normalized = int(value)
    except (TypeError, ValueError) as exc:
        raise SettingsError(format_settings_error(f"{name} must be an integer.")) from exc
    if normalized < minimum:
        raise SettingsError(
            format_settings_error(f"{name} must be greater than or equal to {minimum}.")
        )
    return normalized


def normalize_float_setting(module, name, minimum):
    value = getattr(module, name)
    if isinstance(value, bool):
        raise SettingsError(format_settings_error(f"{name} must be a number."))
    try:
        normalized = float(value)
    except (TypeError, ValueError) as exc:
        raise SettingsError(format_settings_error(f"{name} must be a number.")) from exc
    if normalized < minimum:
        raise SettingsError(
            format_settings_error(f"{name} must be greater than or equal to {minimum}.")
        )
    return normalized


def load_settings(path=None):
    settings_path = Path(path or SETTINGS_PATH)
    if not settings_path.exists():
        raise SettingsError(format_settings_error(f"File not found: {settings_path}"))

    spec = importlib.util.spec_from_file_location(
        "_outlook_calendar_sync_user_settings",
        settings_path,
    )
    if spec is None or spec.loader is None:
        raise SettingsError(format_settings_error(f"Could not load: {settings_path}"))

    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        raise SettingsError(format_settings_error(f"Could not read settings.py: {exc}")) from exc

    missing = [name for name in REQUIRED_SETTINGS if not hasattr(module, name)]
    if missing:
        raise SettingsError(
            format_settings_error("Missing required setting(s): " + ", ".join(missing))
        )

    store_a_hint = normalize_text_setting(module, "STORE_A_HINT")
    store_b_hint = normalize_text_setting(module, "STORE_B_HINT")
    if store_a_hint.lower() == store_b_hint.lower():
        raise SettingsError(
            format_settings_error("STORE_A_HINT and STORE_B_HINT must be different.")
        )

    return SyncSettings(
        store_a_hint,
        store_b_hint,
        normalize_int_setting(module, "DAYS_BACK", minimum=0),
        normalize_int_setting(module, "DAYS_FORWARD", minimum=1),
        normalize_int_setting(module, "POLL_SECONDS", minimum=1),
        normalize_int_setting(module, "SCAN_BATCH_SIZE", minimum=0),
        normalize_float_setting(module, "SCAN_BATCH_PAUSE_SECONDS", minimum=0),
        normalize_int_setting(module, "OUTLOOK_READY_TIMEOUT_SECONDS", minimum=1),
        normalize_int_setting(module, "OUTLOOK_READY_RETRY_SECONDS", minimum=1),
    )


class MissingPywin32Dependency(RuntimeError):
    pass


class OutlookStartupTimeout(RuntimeError):
    pass


class AlreadyRunning(RuntimeError):
    pass


class ControlActionInProgress(RuntimeError):
    pass


class SelectedItemSyncError(RuntimeError):
    pass


def format_pywin32_error(exc):
    return (
        "Missing or unusable pywin32 dependency.\n\n"
        "This script uses Classic Outlook COM automation. The pythoncom and "
        "win32com.client modules are installed by the pywin32 package; do not "
        "install pythoncom separately.\n\n"
        "Install it for the same Python interpreter used to run this script:\n"
        f'    "{sys.executable}" -m pip install --user -r requirements.txt\n\n'
        f"Original import error: {exc}"
    )


def load_outlook_com_modules():
    try:
        pythoncom_module = importlib.import_module("pythoncom")
        win32com_client = importlib.import_module("win32com.client")
    except ImportError as exc:
        raise MissingPywin32Dependency(format_pywin32_error(exc)) from exc
    return pythoncom_module, win32com_client


def setup_logging():
    global LOGGING_READY
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=LOG_PATH,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        force=True,
    )
    LOGGING_READY = True


def format_unexpected_error(exc, include_log_path=False):
    message = (
        "OutlookCalendarSync stopped because of an error.\n\n"
        f"{type(exc).__name__}: {exc}"
    )
    if include_log_path:
        message += f"\n\nMore details were written to:\n    {LOG_PATH}"
    return message


def logging_is_configured():
    return LOGGING_READY


def process_is_running(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False

    if os.name == "nt":
        try:
            import ctypes
        except ImportError:
            return False

        synchronize = 0x00100000
        wait_timeout = 0x00000102
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(synchronize, False, pid)
        if not handle:
            return False
        try:
            return kernel32.WaitForSingleObject(handle, 0) == wait_timeout
        finally:
            kernel32.CloseHandle(handle)

    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def unique_pids(values):
    pids = []
    seen = set()
    for value in values:
        try:
            pid = int(value)
        except (TypeError, ValueError):
            continue
        if pid <= 0 or pid in seen:
            continue
        seen.add(pid)
        pids.append(pid)
    return pids


def read_json_dict(path):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}
    if not isinstance(data, dict):
        return {}
    return data


def read_process_registry(process_registry_path=None):
    path = Path(process_registry_path or PROCESS_REGISTRY_PATH)
    data = read_json_dict(path)
    return unique_pids(data.get("pids", []))


def write_process_registry(pids, process_registry_path=None):
    path = Path(process_registry_path or PROCESS_REGISTRY_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "pids": unique_pids(pids),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "script": str(Path(__file__).resolve()),
    }
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return data


def register_process(pid, process_registry_path=None):
    pids = read_process_registry(process_registry_path=process_registry_path)
    pids.append(pid)
    return write_process_registry(pids, process_registry_path=process_registry_path)


def known_process_pids(status_path=None, process_registry_path=None):
    pids = []
    status = read_status(status_path=status_path)
    if status:
        pids.append(status["pid"])
    pids.extend(read_process_registry(process_registry_path=process_registry_path))
    return unique_pids(pids)


def running_known_pids(
    status_path=None,
    process_registry_path=None,
    is_running_func=None,
    exclude_pids=None,
):
    is_running_func = is_running_func or process_is_running
    excluded = set(unique_pids(exclude_pids or []))
    return [
        pid
        for pid in known_process_pids(
            status_path=status_path,
            process_registry_path=process_registry_path,
        )
        if pid not in excluded and is_running_func(pid)
    ]


def cleanup_process_registry(process_registry_path=None, is_running_func=None):
    is_running_func = is_running_func or process_is_running
    live_pids = [
        pid
        for pid in read_process_registry(process_registry_path=process_registry_path)
        if is_running_func(pid)
    ]
    write_process_registry(live_pids, process_registry_path=process_registry_path)
    return live_pids


class ControlLock:
    def __init__(self, path):
        self.path = Path(path)
        self.released = False

    def release(self):
        if self.released:
            return
        remove_if_exists(self.path)
        self.released = True

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc, _tb):
        self.release()
        return False


def acquire_control_lock(
    lock_path=None,
    timeout_seconds=None,
    is_running_func=None,
    sleep_func=None,
    monotonic_func=None,
):
    path = Path(lock_path or CONTROL_LOCK_PATH)
    timeout_seconds = (
        CONTROL_LOCK_TIMEOUT_SECONDS if timeout_seconds is None else timeout_seconds
    )
    is_running_func = is_running_func or process_is_running
    sleep_func = sleep_func or time.sleep
    monotonic_func = monotonic_func or time.monotonic
    deadline = monotonic_func() + timeout_seconds
    path.parent.mkdir(parents=True, exist_ok=True)

    while True:
        try:
            fd = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            owner = read_json_dict(path)
            owner_pid = owner.get("pid")
            if not owner_pid or not is_running_func(owner_pid):
                remove_if_exists(path)
                continue
            remaining = deadline - monotonic_func()
            if timeout_seconds <= 0 or remaining <= 0:
                raise ControlActionInProgress(
                    "Another OutlookCalendarSync control action is already in progress. "
                    "Try again in a few seconds."
                )
            sleep_func(min(CONTROL_SLEEP_SECONDS, remaining))
            continue

        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(
                {
                    "pid": os.getpid(),
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                },
                handle,
                indent=2,
            )
        return ControlLock(path)


def terminate_process(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0 or pid == os.getpid():
        return False

    if os.name == "nt":
        try:
            result = subprocess.run(
                ["taskkill.exe", "/PID", str(pid), "/T", "/F"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=10,
            )
        except (OSError, subprocess.SubprocessError):
            return False
        return result.returncode == 0

    try:
        os.kill(pid, 15)
    except OSError:
        return False
    return True


def append_control_log(message, control_log_path=None):
    path = Path(control_log_path or CONTROL_LOG_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(f"{datetime.now().isoformat(timespec='seconds')} {message}\n")


def read_control_log_tail(control_log_path=None, max_chars=2000):
    path = Path(control_log_path or CONTROL_LOG_PATH)
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""
    return text[-max_chars:].strip()


def read_status(status_path=None):
    path = Path(status_path or STATUS_PATH)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None
    if not isinstance(data, dict):
        return None
    try:
        data["pid"] = int(data.get("pid"))
    except (TypeError, ValueError):
        return None
    return data


def get_running_status(status_path=None, process_registry_path=None, is_running_func=None):
    is_running_func = is_running_func or process_is_running
    status = read_status(status_path=status_path)
    if status is not None and is_running_func(status["pid"]):
        return status
    registered = running_known_pids(
        status_path=None,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    if registered:
        return {
            "pid": registered[0],
            "started_at": "unknown",
            "state": "starting",
            "message": "Process has started and is preparing status.",
        }
    return None


def write_status(status_path=None, state="starting", message="", settings_path=None):
    path = Path(status_path or STATUS_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "pid": os.getpid(),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "script": str(Path(__file__).resolve()),
        "state": state,
    }
    if message:
        data["message"] = str(message)
    if settings_path:
        data["settings_path"] = str(Path(settings_path))
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return data


def update_status_state(state, message="", status_path=None, settings_path=None):
    path = Path(status_path or STATUS_PATH)
    data = read_status(status_path=path) or {
        "pid": os.getpid(),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "script": str(Path(__file__).resolve()),
    }
    data["state"] = state
    if message:
        data["message"] = str(message)
    else:
        data.pop("message", None)
    if settings_path:
        data["settings_path"] = str(Path(settings_path))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return data


def remove_if_exists(path):
    try:
        Path(path).unlink()
    except FileNotFoundError:
        pass


def clear_stop_request(stop_path=None):
    remove_if_exists(stop_path or STOP_REQUEST_PATH)


def clear_running_status(status_path=None, stop_path=None):
    remove_if_exists(status_path or STATUS_PATH)
    clear_stop_request(stop_path=stop_path)


def mark_running(
    status_path=None,
    stop_path=None,
    process_registry_path=None,
    is_running_func=None,
    settings_path=None,
):
    status_path = Path(status_path or STATUS_PATH)
    stop_path = Path(stop_path or STOP_REQUEST_PATH)
    existing_pids = running_known_pids(
        status_path=status_path,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
        exclude_pids=[os.getpid()],
    )
    if existing_pids:
        raise AlreadyRunning(
            f"OutlookCalendarSync is already running with PID {existing_pids[0]}."
        )
    remove_if_exists(status_path)
    clear_stop_request(stop_path=stop_path)
    status = write_status(status_path=status_path, settings_path=settings_path)
    register_process(os.getpid(), process_registry_path=process_registry_path)
    return status


def stop_requested(stop_path=None):
    return Path(stop_path or STOP_REQUEST_PATH).exists()


def request_stop(stop_path=None):
    path = Path(stop_path or STOP_REQUEST_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        datetime.now().isoformat(timespec="seconds"),
        encoding="utf-8",
    )
    return path


def wait_for_stop_or_timeout(seconds, stop_path=None, sleep_func=None, monotonic_func=None):
    sleep_func = sleep_func or time.sleep
    monotonic_func = monotonic_func or time.monotonic
    deadline = monotonic_func() + seconds

    while True:
        if stop_requested(stop_path=stop_path):
            return True
        remaining = deadline - monotonic_func()
        if remaining <= 0:
            return False
        sleep_func(min(CONTROL_SLEEP_SECONDS, remaining))


def print_status(
    status_path=None,
    process_registry_path=None,
    is_running_func=None,
    control_log_path=None,
):
    status = get_running_status(
        status_path=status_path,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    if not status:
        print("OutlookCalendarSync is not running.")
        control_log = read_control_log_tail(control_log_path=control_log_path)
        if control_log:
            print("\nLast control output:")
            print(control_log)
        return 1

    started_at = status.get("started_at", "unknown")
    state = status.get("state", "running")
    print(
        f"OutlookCalendarSync is {state}. "
        f"PID: {status['pid']}. Started: {started_at}."
    )
    settings_path = status.get("settings_path")
    if settings_path:
        print(f"Settings: {settings_path}")
    message = status.get("message")
    if message:
        print(f"Last message: {message}")
    return 0


def start_sync(
    script_path=None,
    python_executable=None,
    settings_path=None,
    status_path=None,
    stop_path=None,
    control_log_path=None,
    process_registry_path=None,
    lock_path=None,
    lock_timeout_seconds=None,
    is_running_func=None,
    use_lock=True,
):
    if use_lock:
        try:
            with acquire_control_lock(
                lock_path=lock_path,
                timeout_seconds=lock_timeout_seconds,
                is_running_func=is_running_func,
            ):
                return start_sync(
                    script_path=script_path,
                    python_executable=python_executable,
                    settings_path=settings_path,
                    status_path=status_path,
                    stop_path=stop_path,
                    control_log_path=control_log_path,
                    process_registry_path=process_registry_path,
                    lock_path=lock_path,
                    lock_timeout_seconds=lock_timeout_seconds,
                    is_running_func=is_running_func,
                    use_lock=False,
                )
        except ControlActionInProgress as exc:
            print(exc)
            return 1

    script_path = Path(script_path or Path(__file__).resolve())
    python_executable = python_executable or sys.executable

    running_pids = running_known_pids(
        status_path=status_path,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    if running_pids:
        print(f"Already running. PID: {running_pids[0]}.")
        cleanup_process_registry(
            process_registry_path=process_registry_path,
            is_running_func=is_running_func,
        )
        return 0

    cleanup_process_registry(
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    clear_stop_request(stop_path=stop_path)
    control_log_path = Path(control_log_path or CONTROL_LOG_PATH)
    control_log_path.parent.mkdir(parents=True, exist_ok=True)
    command = [python_executable, str(script_path), "--run"]
    if settings_path:
        command.extend(["--settings", str(Path(settings_path))])
    creationflags = 0
    if os.name == "nt":
        creationflags = (
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )

    append_control_log("Start requested.", control_log_path=control_log_path)
    with control_log_path.open("ab") as control_log:
        process = subprocess.Popen(
            command,
            cwd=str(script_path.parent),
            stdin=subprocess.DEVNULL,
            stdout=control_log,
            stderr=control_log,
            creationflags=creationflags,
        )
    register_process(process.pid, process_registry_path=process_registry_path)
    print(f"Start requested for OutlookCalendarSync. PID: {process.pid}.")
    return 0


def stop_sync(
    status_path=None,
    stop_path=None,
    process_registry_path=None,
    lock_path=None,
    lock_timeout_seconds=None,
    is_running_func=None,
    terminate_func=None,
    wait_seconds=STOP_WAIT_SECONDS,
    sleep_func=None,
    monotonic_func=None,
    use_lock=True,
):
    if use_lock:
        try:
            with acquire_control_lock(
                lock_path=lock_path,
                timeout_seconds=lock_timeout_seconds,
                is_running_func=is_running_func,
                sleep_func=sleep_func,
                monotonic_func=monotonic_func,
            ):
                return stop_sync(
                    status_path=status_path,
                    stop_path=stop_path,
                    process_registry_path=process_registry_path,
                    lock_path=lock_path,
                    lock_timeout_seconds=lock_timeout_seconds,
                    is_running_func=is_running_func,
                    terminate_func=terminate_func,
                    wait_seconds=wait_seconds,
                    sleep_func=sleep_func,
                    monotonic_func=monotonic_func,
                    use_lock=False,
                )
        except ControlActionInProgress as exc:
            print(exc)
            return 1

    is_running_func = is_running_func or process_is_running
    terminate_func = terminate_func or terminate_process
    pids = running_known_pids(
        status_path=status_path,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    if not pids:
        cleanup_process_registry(
            process_registry_path=process_registry_path,
            is_running_func=is_running_func,
        )
        print("OutlookCalendarSync is not running.")
        return 0

    request_stop(stop_path=stop_path)
    if read_status(status_path=status_path):
        update_status_state("stopping", status_path=status_path)
    print("Stop requested for OutlookCalendarSync. PID(s): " + ", ".join(map(str, pids)) + ".")

    if wait_seconds > 0:
        sleep_func = sleep_func or time.sleep
        monotonic_func = monotonic_func or time.monotonic
        deadline = monotonic_func() + wait_seconds
        while monotonic_func() < deadline:
            remaining = running_known_pids(
                status_path=status_path,
                process_registry_path=process_registry_path,
                is_running_func=is_running_func,
            )
            if not remaining:
                remove_if_exists(status_path or STATUS_PATH)
                write_process_registry([], process_registry_path=process_registry_path)
                print("Stopped OutlookCalendarSync.")
                return 0
            sleep_func(CONTROL_SLEEP_SECONDS)

    remaining = running_known_pids(
        status_path=status_path,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    for pid in remaining:
        terminate_func(pid)

    still_running = running_known_pids(
        status_path=status_path,
        process_registry_path=process_registry_path,
        is_running_func=is_running_func,
    )
    if still_running:
        write_process_registry(still_running, process_registry_path=process_registry_path)
        print(
            "Stop requested, but these OutlookCalendarSync PID(s) are still running: "
            + ", ".join(map(str, still_running))
            + "."
        )
        return 1

    remove_if_exists(status_path or STATUS_PATH)
    write_process_registry([], process_registry_path=process_registry_path)
    print("Stopped OutlookCalendarSync. PID(s): " + ", ".join(map(str, remaining)) + ".")
    return 0


def restart_sync(
    script_path=None,
    python_executable=None,
    settings_path=None,
    status_path=None,
    stop_path=None,
    control_log_path=None,
    process_registry_path=None,
    lock_path=None,
    lock_timeout_seconds=None,
    is_running_func=None,
    terminate_func=None,
    wait_seconds=STOP_WAIT_SECONDS,
    sleep_func=None,
    monotonic_func=None,
):
    try:
        with acquire_control_lock(
            lock_path=lock_path,
            timeout_seconds=lock_timeout_seconds,
            is_running_func=is_running_func,
            sleep_func=sleep_func,
            monotonic_func=monotonic_func,
        ):
            running_pids = running_known_pids(
                status_path=status_path,
                process_registry_path=process_registry_path,
                is_running_func=is_running_func,
            )
            if not running_pids:
                cleanup_process_registry(
                    process_registry_path=process_registry_path,
                    is_running_func=is_running_func,
                )
                print("OutlookCalendarSync is not running. Use Start to start it first.")
                return 1

            stop_result = stop_sync(
                status_path=status_path,
                stop_path=stop_path,
                process_registry_path=process_registry_path,
                is_running_func=is_running_func,
                terminate_func=terminate_func,
                wait_seconds=wait_seconds,
                sleep_func=sleep_func,
                monotonic_func=monotonic_func,
                use_lock=False,
            )
            if stop_result != 0:
                return stop_result

            start_result = start_sync(
                script_path=script_path,
                python_executable=python_executable,
                settings_path=settings_path,
                status_path=status_path,
                stop_path=stop_path,
                control_log_path=control_log_path,
                process_registry_path=process_registry_path,
                is_running_func=is_running_func,
                use_lock=False,
            )
            if start_result == 0:
                print("Restarted OutlookCalendarSync.")
            return start_result
    except ControlActionInProgress as exc:
        print(exc)
        return 1


def canonical_datetime(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def safe_get(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def calendar_folder_path(calendar):
    return safe_get(calendar, "FolderPath", "") or ""


def call_outlook_member(obj, name, default=None):
    if obj is None:
        return default
    try:
        value = getattr(obj, name)
        if callable(value):
            return value()
        return value
    except Exception:
        return default


def is_appointment_item(item):
    return safe_get(item, "Class") == 26


def selected_appointment_from_outlook(outlook_app):
    inspector = call_outlook_member(outlook_app, "ActiveInspector")
    current_item = safe_get(inspector, "CurrentItem")
    if is_appointment_item(current_item):
        return current_item

    explorer = call_outlook_member(outlook_app, "ActiveExplorer")
    selection = safe_get(explorer, "Selection")
    count = safe_get(selection, "Count", 0) or 0
    try:
        count = int(count)
    except (TypeError, ValueError):
        count = 0

    if count == 0:
        raise SelectedItemSyncError("Select or open one calendar appointment first.")
    if count > 1:
        raise SelectedItemSyncError("Select only one calendar appointment.")

    try:
        item = selection.Item(1)
    except Exception as exc:
        raise SelectedItemSyncError(
            "Could not read the selected Outlook item."
        ) from exc

    if not is_appointment_item(item):
        raise SelectedItemSyncError("The selected Outlook item is not an appointment.")
    return item


def normalized_body(item):
    body = safe_get(item, "Body", "")
    return body or ""


def normalized_categories(item):
    c = safe_get(item, "Categories", "")
    return c or ""


def payload_from_item(item):
    return {
        "Subject": safe_get(item, "Subject", "") or "",
        "Body": normalized_body(item),
        "Start": canonical_datetime(item.Start),
        "End": canonical_datetime(item.End),
        "AllDayEvent": bool(safe_get(item, "AllDayEvent", False)),
        "Location": safe_get(item, "Location", "") or "",
        "Categories": normalized_categories(item),
        "Sensitivity": int(safe_get(item, "Sensitivity", 0) or 0),
        "BusyStatus": int(safe_get(item, "BusyStatus", 2) or 2),
        "Importance": int(safe_get(item, "Importance", 1) or 1),
        "ReminderSet": bool(safe_get(item, "ReminderSet", False)),
        "ReminderMinutesBeforeStart": int(
            safe_get(item, "ReminderMinutesBeforeStart", 15) or 15
        ),
    }


def signature(item):
    raw = json.dumps(
        payload_from_item(item),
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def modification_time(item):
    dt = safe_get(item, "LastModificationTime")
    return canonical_datetime(dt) if dt else ""


def change_snapshot(item, old_sig, old_modified):
    """
    Use Outlook's cheap modification stamp before reading full appointment
    payloads. Loading Body on every poll can make Classic Outlook sluggish.
    """
    current_modified = modification_time(item)
    if old_modified and current_modified == old_modified:
        return False, old_sig, current_modified

    current_sig = signature(item)
    changed = old_sig is not None and current_sig != old_sig
    return changed, current_sig, current_modified


def property_accessor_get(item, prop_name):
    try:
        schema = SYNC_NS + prop_name
        return item.PropertyAccessor.GetProperty(schema)
    except Exception:
        return None


def property_accessor_set(item, prop_name, value):
    schema = SYNC_NS + prop_name
    item.PropertyAccessor.SetProperty(schema, value)


def get_global_identity(item):
    for prop in ("GlobalAppointmentID", "EntryID"):
        val = safe_get(item, prop)
        if val:
            return str(val)
    return ""


def occurrence_key(item):
    """
    Build a logical identity for a visible occurrence.

    For non-recurring items, EntryID/GlobalAppointmentID is enough.

    For recurring occurrences/exceptions, combine the global series identity
    with OriginalDate when available. This keeps a moved exception tied to
    its original slot in the series.
    """
    global_id = get_global_identity(item)

    recurrence_state = safe_get(item, "RecurrenceState", OL_RECURRENCE_STATE_NOT_RECURRING)

    if recurrence_state == OL_RECURRENCE_STATE_NOT_RECURRING:
        return f"single|{global_id}"

    original = safe_get(item, "OriginalDate")
    if original:
        original_key = canonical_datetime(original)
    else:
        original_key = canonical_datetime(item.Start)

    return f"recur|{global_id}|{original_key}"


class Store:
    def __init__(self):
        BASE_DIR.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(DB_PATH)
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute(
            """
            CREATE TABLE IF NOT EXISTS pairs (
                logical_key TEXT PRIMARY KEY,
                a_entry_id TEXT,
                a_store_id TEXT,
                b_entry_id TEXT,
                b_store_id TEXT,
                a_sig TEXT,
                b_sig TEXT,
                a_modified TEXT,
                b_modified TEXT
            )
            """
        )
        self.db.commit()

    def rows(self):
        return list(
            self.db.execute(
                """
                SELECT logical_key,a_entry_id,a_store_id,b_entry_id,b_store_id,
                       a_sig,b_sig,a_modified,b_modified
                FROM pairs
                """
            )
        )

    def upsert(
        self,
        logical_key,
        a_entry_id,
        a_store_id,
        b_entry_id,
        b_store_id,
        a_sig,
        b_sig,
        a_modified,
        b_modified,
    ):
        self.db.execute(
            """
            INSERT INTO pairs(
                logical_key,a_entry_id,a_store_id,b_entry_id,b_store_id,
                a_sig,b_sig,a_modified,b_modified
            )
            VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(logical_key) DO UPDATE SET
                a_entry_id=excluded.a_entry_id,
                a_store_id=excluded.a_store_id,
                b_entry_id=excluded.b_entry_id,
                b_store_id=excluded.b_store_id,
                a_sig=excluded.a_sig,
                b_sig=excluded.b_sig,
                a_modified=excluded.a_modified,
                b_modified=excluded.b_modified
            """,
            (
                logical_key,
                a_entry_id,
                a_store_id,
                b_entry_id,
                b_store_id,
                a_sig,
                b_sig,
                a_modified,
                b_modified,
            ),
        )
        self.db.commit()

    def delete(self, logical_key):
        self.db.execute("DELETE FROM pairs WHERE logical_key=?", (logical_key,))
        self.db.commit()


class OutlookSync:
    def __init__(self, settings=None, com_modules=None):
        if settings is None:
            settings = load_settings()
        if com_modules is None:
            com_modules = load_outlook_com_modules()

        self.settings = settings
        pythoncom_module, win32com_client = com_modules
        pythoncom_module.CoInitialize()
        self.outlook = win32com_client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.store = Store()

        self.calendar_a = self.find_calendar(self.settings.STORE_A_HINT)
        self.calendar_b = self.find_calendar(self.settings.STORE_B_HINT)

        logging.info(
            "Calendar A: %s | Calendar B: %s",
            self.calendar_a.FolderPath,
            self.calendar_b.FolderPath,
        )

    def find_calendar(self, hint):
        hint_lower = hint.lower()
        candidates = []

        for i in range(1, self.ns.Stores.Count + 1):
            st = self.ns.Stores.Item(i)
            display = str(safe_get(st, "DisplayName", "") or "")
            root = st.GetRootFolder()
            root_name = str(safe_get(root, "Name", "") or "")
            text = f"{display} {root_name}".lower()
            candidates.append(display)

            if hint_lower in text:
                try:
                    return st.GetDefaultFolder(OL_FOLDER_CALENDAR)
                except Exception:
                    pass

        raise RuntimeError(
            f"Could not find Outlook store matching {hint!r}. "
            f"Available stores: {candidates}"
        )

    def expanded_items(self, calendar):
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = datetime.now() - timedelta(days=self.settings.DAYS_BACK)
        end = datetime.now() + timedelta(days=self.settings.DAYS_FORWARD)

        restriction = (
            "[Start] <= '"
            + end.strftime("%m/%d/%Y %I:%M %p")
            + "' AND [End] >= '"
            + start.strftime("%m/%d/%Y %I:%M %p")
            + "'"
        )

        restricted = items.Restrict(restriction)

        result = {}
        scanned = 0
        item = restricted.GetFirst()
        while item is not None:
            scanned += 1
            try:
                if safe_get(item, "Class") != 26:  # olAppointment
                    item = restricted.GetNext()
                    continue

                # Mirrored standalone appointments carry the logical key of
                # their source occurrence. Native items derive their key from
                # Outlook recurrence identity. This is what makes both halves
                # of a pair appear under the same key during reconciliation.
                marker = property_accessor_get(item, SYNC_PROP)
                key = str(marker) if marker else occurrence_key(item)
                result[key] = item

                if (
                    self.settings.SCAN_BATCH_SIZE > 0
                    and self.settings.SCAN_BATCH_PAUSE_SECONDS > 0
                    and scanned % self.settings.SCAN_BATCH_SIZE == 0
                ):
                    time.sleep(self.settings.SCAN_BATCH_PAUSE_SECONDS)
            except Exception:
                logging.exception("Failed reading calendar item")
            item = restricted.GetNext()
        return result

    def copy_into_calendar(self, source, target_calendar, logical_key):
        target = target_calendar.Items.Add(OL_APPOINTMENT_ITEM)
        self.apply_payload(source, target)
        property_accessor_set(target, SYNC_PROP, logical_key)
        target.Save()
        return target

    def apply_payload(self, source, target):
        p = payload_from_item(source)

        target.Subject = p["Subject"]
        target.Body = p["Body"]
        target.Start = source.Start
        target.End = source.End
        target.AllDayEvent = p["AllDayEvent"]
        target.Location = p["Location"]
        target.Categories = p["Categories"]
        target.Sensitivity = p["Sensitivity"]
        target.BusyStatus = p["BusyStatus"]
        target.Importance = p["Importance"]
        target.ReminderSet = p["ReminderSet"]
        if p["ReminderSet"]:
            target.ReminderMinutesBeforeStart = p["ReminderMinutesBeforeStart"]

    def get_item(self, entry_id, store_id):
        if not entry_id:
            return None
        try:
            return self.ns.GetItemFromID(entry_id, store_id)
        except Exception:
            return None

    def delete_item(self, item):
        if item is None:
            return
        try:
            item.Delete()
        except Exception:
            logging.exception("Delete failed")

    def selected_item(self):
        return selected_appointment_from_outlook(self.outlook)

    def item_calendar_side(self, item):
        parent = safe_get(item, "Parent")
        store_id = safe_get(parent, "StoreID")
        folder_path = calendar_folder_path(parent).lower()
        calendar_a_path = calendar_folder_path(self.calendar_a).lower()
        calendar_b_path = calendar_folder_path(self.calendar_b).lower()

        if folder_path and calendar_a_path and folder_path == calendar_a_path:
            return "a"
        if folder_path and calendar_b_path and folder_path == calendar_b_path:
            return "b"
        if not folder_path and store_id and store_id == safe_get(self.calendar_a, "StoreID"):
            return "a"
        if not folder_path and store_id and store_id == safe_get(self.calendar_b, "StoreID"):
            return "b"
        return None

    def row_for_key(self, logical_key):
        for row in self.store.rows():
            if row[0] == logical_key:
                return row
        return None

    def item_logical_key(self, item):
        marker = property_accessor_get(item, SYNC_PROP)
        return str(marker) if marker else occurrence_key(item)

    def target_from_row(self, row, side):
        if row is None:
            return None
        if side == "a":
            return self.get_item(row[1], row[2])
        return self.get_item(row[3], row[4])

    def upsert_selected_pair(self, logical_key, side, source, target):
        if side == "a":
            a_item = source
            b_item = target
        else:
            a_item = target
            b_item = source

        self.store.upsert(
            logical_key,
            safe_get(a_item, "EntryID"),
            safe_get(safe_get(a_item, "Parent"), "StoreID"),
            safe_get(b_item, "EntryID"),
            safe_get(safe_get(b_item, "Parent"), "StoreID"),
            signature(a_item),
            signature(b_item),
            modification_time(a_item),
            modification_time(b_item),
        )

    def sync_selected_item(self, item=None):
        item = item or self.selected_item()
        if not is_appointment_item(item):
            raise SelectedItemSyncError("The selected Outlook item is not an appointment.")

        source_side = self.item_calendar_side(item)
        if source_side is None:
            raise SelectedItemSyncError(
                "The selected appointment is not in one of the configured calendars."
            )

        logical_key = self.item_logical_key(item)
        row = self.row_for_key(logical_key)
        target_side = "b" if source_side == "a" else "a"
        target_calendar = self.calendar_b if target_side == "b" else self.calendar_a
        target_label = "Calendar B" if target_side == "b" else "Calendar A"
        target = self.target_from_row(row, target_side)

        if target is None:
            target = self.copy_into_calendar(item, target_calendar, logical_key)
            action = "Created"
        else:
            self.apply_payload(item, target)
            target.Save()
            action = "Updated"

        self.upsert_selected_pair(logical_key, source_side, item, target)
        return f"{action} copy in {target_label}: {safe_get(item, 'Subject', '') or '(no subject)'}"

    def reconcile(self):
        a_view = self.expanded_items(self.calendar_a)
        b_view = self.expanded_items(self.calendar_b)
        rows = {r[0]: r for r in self.store.rows()}
        deleted_keys = set()

        # Existing logical pairs
        for key, row in list(rows.items()):
            (
                _,
                a_entry,
                a_store,
                b_entry,
                b_store,
                old_a_sig,
                old_b_sig,
                old_a_mod,
                old_b_mod,
            ) = row

            ae = a_view.get(key)
            be = b_view.get(key)

            # expanded_items() indexes mirrored copies by the source logical
            # key stored in their custom MAPI property. If a key is absent,
            # that side's item/occurrence has genuinely disappeared from the
            # configured synchronization window.

            if ae is None and be is None:
                self.store.delete(key)
                deleted_keys.add(key)
                continue

            if ae is None:
                logging.info("A deleted %s; deleting B twin", key)
                self.delete_item(be)
                self.store.delete(key)
                deleted_keys.add(key)
                continue

            if be is None:
                logging.info("B deleted %s; deleting A twin", key)
                self.delete_item(ae)
                self.store.delete(key)
                deleted_keys.add(key)
                continue

            a_changed, a_sig, a_mod = change_snapshot(ae, old_a_sig, old_a_mod)
            b_changed, b_sig, b_mod = change_snapshot(be, old_b_sig, old_b_mod)

            if not a_changed and not b_changed:
                if a_mod != old_a_mod or b_mod != old_b_mod:
                    self.store.upsert(
                        key,
                        a_entry,
                        a_store,
                        b_entry,
                        b_store,
                        a_sig,
                        b_sig,
                        a_mod,
                        b_mod,
                    )
                continue

            if a_changed and not b_changed:
                logging.info("Updating B from A: %s", key)
                self.apply_payload(ae, be)
                be.Save()
                b_sig = a_sig
            elif b_changed and not a_changed:
                logging.info("Updating A from B: %s", key)
                self.apply_payload(be, ae)
                ae.Save()
                a_sig = b_sig
            else:
                if a_mod >= b_mod:
                    logging.warning("Conflict; A wins: %s", key)
                    self.apply_payload(ae, be)
                    be.Save()
                    b_sig = a_sig
                else:
                    logging.warning("Conflict; B wins: %s", key)
                    self.apply_payload(be, ae)
                    ae.Save()
                    a_sig = b_sig

            self.store.upsert(
                key,
                safe_get(ae, "EntryID"),
                safe_get(ae, "Parent").StoreID,
                safe_get(be, "EntryID"),
                safe_get(be, "Parent").StoreID,
                a_sig,
                b_sig,
                modification_time(ae),
                modification_time(be),
            )

        rows = {r[0]: r for r in self.store.rows()}

        # New native A occurrences
        created_b_twins = False
        for key, ae in list(a_view.items()):
            if key in deleted_keys:
                continue
            if key in rows:
                continue
            if property_accessor_get(ae, SYNC_PROP):
                continue

            logging.info("Creating B twin: %s", key)
            be = self.copy_into_calendar(ae, self.calendar_b, key)
            created_b_twins = True

            self.store.upsert(
                key,
                safe_get(ae, "EntryID"),
                safe_get(ae, "Parent").StoreID,
                safe_get(be, "EntryID"),
                safe_get(be, "Parent").StoreID,
                signature(ae),
                signature(be),
                modification_time(ae),
                modification_time(be),
            )

        rows = {r[0]: r for r in self.store.rows()}

        if created_b_twins:
            # Re-read B only when A->B creations changed the view.
            b_view = self.expanded_items(self.calendar_b)

        # New native B occurrences
        for key, be in list(b_view.items()):
            if key in deleted_keys:
                continue
            if key in rows:
                continue
            if property_accessor_get(be, SYNC_PROP):
                continue

            logging.info("Creating A twin: %s", key)
            ae = self.copy_into_calendar(be, self.calendar_a, key)

            self.store.upsert(
                key,
                safe_get(ae, "EntryID"),
                safe_get(ae, "Parent").StoreID,
                safe_get(be, "EntryID"),
                safe_get(be, "Parent").StoreID,
                signature(ae),
                signature(be),
                modification_time(ae),
                modification_time(be),
            )

    def run(self):
        logging.info("OutlookCalendarSync started")
        while True:
            if stop_requested():
                break
            try:
                with acquire_control_lock(
                    lock_path=SYNC_LOCK_PATH,
                    timeout_seconds=SYNC_LOCK_TIMEOUT_SECONDS,
                ):
                    self.reconcile()
            except ControlActionInProgress as exc:
                if logging_is_configured():
                    logging.info("Skipping sync pass: %s", exc)
            except Exception as exc:
                if logging_is_configured():
                    logging.exception("Reconciliation failed")
                print(
                    format_unexpected_error(
                        exc,
                        include_log_path=logging_is_configured(),
                    ),
                    file=sys.stderr,
                    flush=True,
                )
            if wait_for_stop_or_timeout(self.settings.POLL_SECONDS):
                break
        logging.info("OutlookCalendarSync stopped")
        clear_stop_request()
        print("OutlookCalendarSync stopped.", flush=True)


def format_wait_seconds(seconds):
    if seconds % 60 == 0:
        minutes = seconds // 60
        label = "minute" if minutes == 1 else "minutes"
        return f"{minutes} {label}"
    return f"{seconds} seconds"


def format_outlook_startup_timeout(settings, timeout_seconds, last_error):
    return (
        "Classic Outlook and the configured calendars were not ready after "
        f"{format_wait_seconds(timeout_seconds)}.\n\n"
        "Open Classic Outlook, confirm both accounts are signed in, and confirm "
        "both calendars appear in the Outlook folder pane.\n\n"
        "Configured account/store hints:\n"
        f"    STORE_A_HINT = {settings.STORE_A_HINT!r}\n"
        f"    STORE_B_HINT = {settings.STORE_B_HINT!r}\n\n"
        f"Last error: {last_error}"
    )


def wait_for_outlook_ready(
    settings,
    com_modules,
    settings_path=None,
    status_path=None,
    stop_path=None,
    sync_factory=None,
    sleep_func=None,
    monotonic_func=None,
):
    sync_factory = sync_factory or OutlookSync
    sleep_func = sleep_func or time.sleep
    monotonic_func = monotonic_func or time.monotonic

    timeout_seconds = settings.OUTLOOK_READY_TIMEOUT_SECONDS
    retry_seconds = settings.OUTLOOK_READY_RETRY_SECONDS
    deadline = monotonic_func() + timeout_seconds

    while True:
        if stop_path is not None and stop_requested(stop_path=stop_path):
            raise KeyboardInterrupt

        try:
            return sync_factory(settings=settings, com_modules=com_modules)
        except Exception as exc:
            now = monotonic_func()
            if status_path is not None:
                update_status_state(
                    "waiting",
                    message=exc,
                    status_path=status_path,
                    settings_path=settings_path,
                )
            if now >= deadline:
                raise OutlookStartupTimeout(
                    format_outlook_startup_timeout(settings, timeout_seconds, exc)
                ) from exc

            wait_seconds = min(retry_seconds, deadline - now)
            logging.info(
                "Waiting for Classic Outlook/calendars to become ready: %s",
                exc,
            )
            if stop_path is None:
                sleep_func(wait_seconds)
            elif wait_for_stop_or_timeout(
                wait_seconds,
                stop_path=stop_path,
                sleep_func=sleep_func,
                monotonic_func=monotonic_func,
            ):
                raise KeyboardInterrupt


def install_autostart(settings_path=None):
    STARTUP_DIR.mkdir(parents=True, exist_ok=True)
    script_path = Path(__file__).resolve()
    python_exe = Path(sys.executable).resolve()
    command = f'start "" /min "{python_exe}" "{script_path}" --start'
    if settings_path:
        command += f' --settings "{Path(settings_path)}"'

    content = (
        "@echo off\r\n"
        f'cd /d "{script_path.parent}"\r\n'
        f"{command}\r\n"
    )
    STARTUP_CMD.write_text(content, encoding="utf-8")
    print(f"Installed per-user autostart: {STARTUP_CMD}")


def remove_autostart():
    if STARTUP_CMD.exists():
        STARTUP_CMD.unlink()
        print(f"Removed per-user autostart: {STARTUP_CMD}")
    else:
        print("Autostart entry was not present.")


def sync_selected_item_once(settings_path=None):
    settings_path = settings_path or SETTINGS_PATH
    settings = load_settings(settings_path)
    com_modules = load_outlook_com_modules()
    setup_logging()
    sync = OutlookSync(settings=settings, com_modules=com_modules)
    selected = sync.selected_item()

    try:
        with acquire_control_lock(
            lock_path=SYNC_LOCK_PATH,
            timeout_seconds=SYNC_LOCK_TIMEOUT_SECONDS,
        ):
            return sync.sync_selected_item(selected)
    except ControlActionInProgress as exc:
        raise SelectedItemSyncError(str(exc)) from exc


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--install-autostart", action="store_true")
    parser.add_argument("--remove-autostart", action="store_true")
    parser.add_argument("--start", action="store_true")
    parser.add_argument("--stop", action="store_true")
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--restart", action="store_true")
    parser.add_argument("--sync-selected", action="store_true")
    parser.add_argument("--run", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--settings", type=Path)
    args = parser.parse_args()

    logging_ready = False
    running_marked = False

    try:
        control_count = sum(
            bool(value)
            for value in (
                args.install_autostart,
                args.remove_autostart,
                args.start,
                args.stop,
                args.status,
                args.restart,
                args.sync_selected,
            )
        )
        if control_count > 1:
            parser.error("choose only one command option")

        if args.install_autostart:
            install_autostart(settings_path=args.settings)
            return 0

        if args.remove_autostart:
            remove_autostart()
            return 0

        if args.start:
            return start_sync(settings_path=args.settings)

        if args.stop:
            return stop_sync()

        if args.status:
            return print_status()

        if args.restart:
            return restart_sync(settings_path=args.settings)

        if args.sync_selected:
            print(sync_selected_item_once(settings_path=args.settings))
            return 0

        settings_path = args.settings or SETTINGS_PATH
        settings = load_settings(settings_path)

        com_modules = load_outlook_com_modules()

        setup_logging()
        logging_ready = True

        try:
            mark_running(settings_path=settings_path)
            running_marked = True
        except AlreadyRunning as exc:
            print(exc)
            return 0

        update_status_state(
            "starting",
            message="Connecting to Classic Outlook and calendars.",
            settings_path=settings_path,
        )
        sync = wait_for_outlook_ready(
            settings=settings,
            com_modules=com_modules,
            settings_path=settings_path,
            status_path=STATUS_PATH,
            stop_path=STOP_REQUEST_PATH,
        )
        update_status_state(
            "syncing",
            message=(
                f"Calendar A: {calendar_folder_path(sync.calendar_a)} | "
                f"Calendar B: {calendar_folder_path(sync.calendar_b)}"
            ),
            settings_path=settings_path,
        )

        sync.run()
        return 0
    except (SettingsError, MissingPywin32Dependency, SelectedItemSyncError) as exc:
        print(exc, file=sys.stderr)
        return 1
    except OutlookStartupTimeout as exc:
        if logging_ready:
            logging.error("%s", exc)
        print(exc, file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("OutlookCalendarSync stopped by user.", file=sys.stderr)
        return 130
    except Exception as exc:
        if logging_ready:
            logging.exception("OutlookCalendarSync stopped because of an error")
        print(
            format_unexpected_error(exc, include_log_path=logging_ready),
            file=sys.stderr,
        )
        return 1
    finally:
        if running_marked:
            clear_running_status()


if __name__ == "__main__":
    raise SystemExit(main())
