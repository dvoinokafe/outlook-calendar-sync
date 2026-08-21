# Outlook Calendar Sync

A no-admin, per-user synchronizer for two Microsoft 365 calendars that are already configured in the same **Classic Outlook for Windows** profile.

It is designed for users who need to keep calendars from two different companies aligned but cannot obtain tenant-admin consent for Microsoft Graph and cannot install a Windows Service.

## Features

- Bidirectional synchronization between two Outlook calendar stores
- Create, update, and delete propagation
- Recurring-series support
- Individual recurring occurrence changes and deletions
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
- `pywin32`

New Outlook is not supported because it does not expose the Classic Outlook COM object model used by this project.

## Installation

Install the dependency for your current Windows user:

```powershell
py -m pip install --user -r requirements.txt
```

Edit these two settings near the top of `outlook_calendar_sync.py`:

```python
STORE_A_HINT = "me@company-a.com"
STORE_B_HINT = "me@company-b.com"
```

Use strings that uniquely identify the two Outlook mailbox/store names.

## Test interactively

Run:

```powershell
py outlook_calendar_sync.py
```

Then test both directions:

1. Create a normal appointment in Calendar A and verify it appears in Calendar B.
2. Modify it and verify the copy changes.
3. Delete it and verify the copy disappears.
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

It intentionally does **not** copy:

- Organizer
- Attendees / recipients
- Meeting-response state
- Online-meeting metadata

This avoids having the mirrored appointment act as a second meeting organizer and prevents duplicate invitations or updates.

## Recurring meetings

Recurring events are expanded inside a rolling calendar window and mirrored occurrence-by-occurrence.

This is deliberate. It allows a single changed or deleted instance of a recurring series to be synchronized independently across unrelated company mailboxes.

Default window:

- 30 days in the past
- 365 days in the future

These values can be changed in the script.

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

Run Outlook and confirm both accounts appear in the folder pane. Then make `STORE_A_HINT` and `STORE_B_HINT` match distinctive portions of the store display names.

### `No module named win32com`

Install the dependency:

```powershell
py -m pip install --user pywin32
```

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
