# Outlook Calendar Sync

A no-admin, per-user synchronizer for two Microsoft 365 calendars that are already configured in the same **Classic Outlook for Windows** profile.

It is designed for users who need to keep calendars from two different companies aligned but cannot obtain tenant-admin consent for Microsoft Graph and cannot install a Windows Service.

## Features

- Bidirectional synchronization between two Outlook calendar stores
- Create, update, and delete propagation
- Delete prompt for synced appointments while the background sync is running
- Recurring-series support
- Individual recurring occurrence changes and deletions
- Mirrored copies marked with the gray `DY Sync` Outlook category
- Local SQLite mapping database to avoid duplicate copies
- Last-writer-wins conflict handling
- Per-user autostart without administrator rights
- No Microsoft Graph application registration
- No tenant admin consent
- No local administrator rights required

## Requirements

- Windows
- Python 3.10+
- Classic Outlook for Windows
- Both Microsoft 365 accounts configured in the same Outlook profile
- `pywin32`, which provides `pythoncom` and `win32com.client`

New Outlook is not supported because it does not expose the Classic Outlook COM object model used by this project.

## Installation

Install the dependency for your current Windows user:

```powershell
py -m pip install --user -r requirements.txt
```

Do not install `pythoncom` separately. It is installed as part of `pywin32`.

Edit these account settings in `settings.py`:

```python
STORE_A_HINT = "me@company-a.com"
STORE_B_HINT = "me@company-b.com"
```

Use strings that uniquely identify the two Outlook mailbox/store names.
The email address is usually the easiest value to match.
Email-address hints also allow the script to recognize invitations sent
between the two configured accounts and skip mirroring them.

You can also tune the scan window and polling periods in `settings.py`:

```python
DAYS_BACK = 2
DAYS_FORWARD = 2
POLL_SECONDS = 5
SCAN_BATCH_SIZE = 25
SCAN_BATCH_PAUSE_SECONDS = 0.05
OUTLOOK_READY_TIMEOUT_SECONDS = 300
OUTLOOK_READY_RETRY_SECONDS = 10
```

## Test interactively

Run:

```powershell
py outlook_calendar_sync.py
```

Then test both directions:

1. Create a normal appointment in Calendar A and verify it appears in Calendar B.
2. Modify it and verify the copy changes.
3. Delete it while the background sync is running and choose whether the
   delete applies only in this account or in both accounts.
4. Repeat from Calendar B.
5. Create a recurring series.
6. Modify only one occurrence and verify only that occurrence changes.
7. Delete only one occurrence and verify only that occurrence disappears.

Logs are written to:

```text
%LOCALAPPDATA%\OutlookCalendarSync\sync.log
```

Synchronization state is stored in:

```text
%LOCALAPPDATA%\OutlookCalendarSync\sync.db
```

Do not delete `sync.db` while using the synchronizer unless you deliberately want to discard all pairing information.

## Start, stop, and status

The script has per-user controls that do not require administrator
rights:

```powershell
py outlook_calendar_sync.py --start
py outlook_calendar_sync.py --stop
py outlook_calendar_sync.py --status
py outlook_calendar_sync.py --restart
py outlook_calendar_sync.py --sync-selected
```

If your real account settings live somewhere else, pass that file
explicitly:

```powershell
py outlook_calendar_sync.py --start --settings "C:\Path\To\settings.py"
py outlook_calendar_sync.py --status
```

`--start` launches the synchronizer in the background if it is not
already running. If a tracked sync process is already alive, it prints
`Already running` and exits without starting a second copy.

`--stop` requests a clean shutdown, waits briefly, then forcibly stops
any tracked sync process that is still alive. This covers a stale copy
that was started by the control command but never reached normal
`status.json` reporting.

`--restart` only restarts when a tracked sync process is already alive.
If the synchronizer is not running, it prints a message asking you to
start it first instead of starting unexpectedly.

Start, stop, and restart are protected by a short control lock so two
button clicks or command-line controls cannot update `status.json`,
`stop.request`, or the process registry at the same time.

`--sync-selected` attaches to the currently open or selected Classic
Outlook appointment and syncs only that item. It works even when the
background synchronizer is not running. If a full sync pass is active,
the selected-item command waits briefly on the sync lock so both paths
do not update Outlook and `sync.db` at the same time.

The selected-item command rejects empty selections, multiple selections,
non-appointment selections, and appointments outside the two configured
calendar stores.

If Outlook reports the selected appointment from an unexpected folder,
the error message includes the selected folder/store, the configured
Calendar A/B folders, and the settings file path being used. Already
synced mirrored copies can still be matched by their sync marker and
local database row when Outlook shows them from a search/result view
instead of the calendar folder itself.

Meetings sent from Calendar A's configured account to Calendar B's
configured account, or from B to A, are skipped by both background sync
and Sync Selected. Outlook already delivers those invitations to the
recipient mailbox, so mirroring them would create duplicates.

`--status` reports whether the script is `starting`, `waiting`, or
`syncing`. If it is waiting for Outlook/calendars, the status output
includes the latest Outlook readiness message so a hidden start does not
look healthy when it is only waiting.

The running status is stored under:

```text
%LOCALAPPDATA%\OutlookCalendarSync\status.json
```

Tracked process IDs are stored under:

```text
%LOCALAPPDATA%\OutlookCalendarSync\processes.json
```

The stop request is stored as:

```text
%LOCALAPPDATA%\OutlookCalendarSync\stop.request
```

Hidden start output and startup errors are written to:

```text
%LOCALAPPDATA%\OutlookCalendarSync\control.log
```

## Start automatically at Windows login

Install a per-user Startup entry:

```powershell
py outlook_calendar_sync.py --install-autostart
```

Remove it with:

```powershell
py outlook_calendar_sync.py --remove-autostart
```

No elevation is required.

This runs at Windows login, not when Outlook itself is opened. If Classic
Outlook or the two configured calendar stores are not ready yet, the
script waits and retries for `OUTLOOK_READY_TIMEOUT_SECONDS` seconds.
With the default settings, it stops after 5 minutes and prints/writes a
message that names the configured account hints and the last Outlook
readiness error.

The Startup entry uses `--start`, so it will not intentionally launch a
second sync loop when the status file or process registry says one is
already running.

## Classic Outlook buttons

`OutlookSyncControls.bas` contains five VBA macros:

- `OutlookSync_Start`
- `OutlookSync_Stop`
- `OutlookSync_Status`
- `OutlookSync_Restart`
- `OutlookSync_SyncSelected`

To use them in Classic Outlook:

1. Open `OutlookSyncControls.bas` and edit `SYNC_SCRIPT` to the full path
   of your `outlook_calendar_sync.py`. The published template leaves this
   blank deliberately so local machine paths are not shared.
2. If your real settings file is not next to the script, edit
   `SETTINGS_FILE` to its full path. Leave it blank to use the adjacent
   `settings.py`.
3. Confirm `PYTHON_COMMAND`. Leave it blank to use `py -3`, set it to a
   specific launcher command such as `py -3.12`, or set it to the full
   path of the `python.exe` that has `pywin32` installed. If you need the
   full path, run `py -0p` in PowerShell and copy the matching Python
   path.
4. In Classic Outlook, open the VBA editor with `Alt+F11`.
5. Import `OutlookSyncControls.bas`.
6. Add the macros to the Quick Access Toolbar or a custom Ribbon group.

The buttons show a small Outlook message box with the actual script
response. Start reports `Already running` when applicable. Stop reports
which tracked sync PID(s) were stopped. Restart refuses to start a new
sync if nothing was already running. Sync Selected copies or updates
only the currently selected/open appointment.

The VBA macro trims configured paths and strips accidental wrapping
quotes, so either `C:\Path\To\outlook_calendar_sync.py` or
`"C:\Path\To\outlook_calendar_sync.py"` can be pasted safely. Embedded
quotes or line breaks are rejected with a clear Outlook message. For
`PYTHON_COMMAND`, use either a simple launcher command like `py -3.12`
or a full `python.exe` path; do not combine a quoted executable path with
extra Python arguments.
Button commands run through a hidden temporary `.cmd` wrapper so paths
with spaces in `SYNC_SCRIPT`, `SETTINGS_FILE`, or `PYTHON_COMMAND` are
handled correctly.

If your company disables Outlook VBA macros, use the command-line
controls above or ask IT whether signed Outlook macros are allowed.

## What is copied

The synchronizer copies:

- Subject
- Body
- Start and end times
- All-day status
- Location
- Categories
- Sensitivity
- Busy status
- Importance
- Reminder settings

Mirrored copies are also marked with the gray Outlook category `DY Sync`.
That marker is reserved for the synchronizer, ignored by the comparison
hash, and not copied back to the original appointment.

It intentionally does **not** copy:

- Organizer
- Attendees / recipients
- Meeting-response state
- Online-meeting metadata

Because organizer and attendee semantics are not copied, direct
invitations between the two configured accounts are left to Outlook's
normal invitation delivery instead of being mirrored by this script.

This avoids having the mirrored appointment act as a second meeting organizer and prevents duplicate invitations or updates.

## Delete choices

When the background synchronizer is running, it listens for deletes from
the two configured calendar folders. If you delete a synced appointment,
Outlook Calendar Sync asks whether to delete it from both configured
accounts, delete it only from the current account, or cancel the delete.

Choosing both accounts lets Outlook delete the selected item, then the
next sync pass deletes the paired copy. Choosing only this account lets
Outlook delete the selected item, then the next sync pass removes only
the local pairing row and keeps the other account's appointment.

This prompt is available only while the background sync process is
running. If the sync process is stopped, Outlook can still delete the
item, and the next reconciliation follows the normal missing-item rule.

## Recurring meetings

Recurring events are expanded inside a rolling calendar window and mirrored occurrence-by-occurrence.

This is deliberate. It allows a single changed or deleted instance of a recurring series to be synchronized independently across unrelated company mailboxes.

Default window:

- 2 days in the past
- 2 days in the future

These values can be changed in `settings.py`.

The default polling interval is 5 seconds. The script also avoids
reading full appointment payloads for unchanged mapped items and pauses
briefly during large scans so Classic Outlook remains more responsive
while you click, open, or move meetings.

## Conflict handling

The synchronizer stores hashes for both sides.

- If only one side changed since the previous reconciliation, that side wins.
- If both sides changed, the appointment with the newer Outlook `LastModificationTime` wins.

## Important limitations

- This runs in the signed-in user's interactive Windows session, not as a Windows Service.
- Classic Outlook must be installed and usable.
- Corporate Outlook security policy may block or prompt for COM automation.
- The sync is limited to the configured rolling date window.
- Moving an occurrence outside that window can temporarily look like a deletion.
- Lower polling intervals or very large recurrence windows can make Outlook sluggish during calendar interaction.
- Organizer and attendee semantics are deliberately not reproduced.
- This project has not been certified by Microsoft and should be tested with non-critical calendar data before production use.

## Security and privacy

This project copies calendar content between two company-controlled mailboxes. That may include confidential information.

Before using it:

- Confirm both employers permit cross-company calendar replication.
- Consider whether appointment subjects, locations, and bodies may contain restricted information.
- Review `SECURITY.md`.
- Use restrictive permissions on `%LOCALAPPDATA%\OutlookCalendarSync`.

## Troubleshooting

### Store not found

Run Outlook and confirm both accounts appear in the folder pane. Then make `STORE_A_HINT` and `STORE_B_HINT` in `settings.py` match distinctive portions of the store display names.

### `Missing or invalid settings.py`

Keep `settings.py` in the same folder as `outlook_calendar_sync.py`. It must define:

```python
STORE_A_HINT
STORE_B_HINT
DAYS_BACK
DAYS_FORWARD
POLL_SECONDS
SCAN_BATCH_SIZE
SCAN_BATCH_PAUSE_SECONDS
OUTLOOK_READY_TIMEOUT_SECONDS
OUTLOOK_READY_RETRY_SECONDS
```

### Outlook is not ready during Windows startup

Open Classic Outlook and confirm both accounts are signed in and both
calendars appear in the folder pane. If Outlook routinely takes longer
than 5 minutes on your machine, increase `OUTLOOK_READY_TIMEOUT_SECONDS`
in `settings.py`.

### `No module named pythoncom` or `No module named win32com`

Install pywin32 for the same Python interpreter used to run the script:

```powershell
py -m pip install --user -r requirements.txt
```

If `PYTHON_COMMAND` is set to a specific launcher version, use the same
version for installation, for example:

```powershell
py -3.12 -m pip install --user -r requirements.txt
```

Then verify the import setup:

```powershell
py -c "import pythoncom, win32com.client"
```

If this still fails, check that the button's `PYTHON_COMMAND` and the
`pip install` command are using the same Python installation or virtual
environment.

### Duplicate appointments

The local database may have been deleted or replaced. Stop the program and inspect:

```text
%LOCALAPPDATA%\OutlookCalendarSync\sync.db
```

### Outlook prompts or blocks automation

This may be controlled by corporate security policy. The project cannot bypass administrative Outlook policy.

## Project status

Early-stage utility. Use at your own risk and test thoroughly before relying on it for business-critical scheduling.

## License

MIT. See [LICENSE](LICENSE).
