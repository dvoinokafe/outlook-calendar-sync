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

    py -m pip install --user pywin32

Python includes sqlite3, json, hashlib, pathlib, and datetime.


CONFIGURATION
=============
Edit only these values:

    STORE_A_HINT
    STORE_B_HINT

Set them to distinctive text that appears in the corresponding Outlook
store/mailbox display names, usually the email address.

Example:

    STORE_A_HINT = "me@company-a.com"
    STORE_B_HINT = "me@company-b.com"

The script searches Outlook stores case-insensitively.

Optional settings:

    DAYS_BACK = 30
    DAYS_FORWARD = 365
    POLL_SECONDS = 30

The rolling window is necessary because Outlook recurrence expansion is
date-window based.


AUTOSTART WITHOUT ADMIN RIGHTS
==============================
Run:

    py outlook_calendar_sync.py --install-autostart

This creates a .cmd launcher in your per-user Windows Startup folder.

Remove it with:

    py outlook_calendar_sync.py --remove-autostart

No administrator rights are required.


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
import json
import logging
import os
import sqlite3
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import pythoncom
import win32com.client

# ---------- USER CONFIGURATION ----------

STORE_A_HINT = "me@company-a.com"
STORE_B_HINT = "me@company-b.com"

DAYS_BACK = 30
DAYS_FORWARD = 365
POLL_SECONDS = 30

# ---------- CONSTANTS ----------

OL_FOLDER_CALENDAR = 9
OL_APPOINTMENT_ITEM = 1
OL_RECURRENCE_STATE_NOT_RECURRING = 0

BASE_DIR = Path(os.environ.get("LOCALAPPDATA", ".")) / "OutlookCalendarSync"
DB_PATH = BASE_DIR / "sync.db"
LOG_PATH = BASE_DIR / "sync.log"

STARTUP_DIR = Path(os.environ.get("APPDATA", ".")) / (
    r"Microsoft\Windows\Start Menu\Programs\Startup"
)
STARTUP_CMD = STARTUP_DIR / "OutlookCalendarSync.cmd"

SYNC_PROP = "OutlookCalendarSyncKey"
SYNC_NS = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/"


def setup_logging():
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=LOG_PATH,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )


def canonical_datetime(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def safe_get(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


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
    def __init__(self):
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.store = Store()

        self.calendar_a = self.find_calendar(STORE_A_HINT)
        self.calendar_b = self.find_calendar(STORE_B_HINT)

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

        start = datetime.now() - timedelta(days=DAYS_BACK)
        end = datetime.now() + timedelta(days=DAYS_FORWARD)

        restriction = (
            "[Start] <= '"
            + end.strftime("%m/%d/%Y %I:%M %p")
            + "' AND [End] >= '"
            + start.strftime("%m/%d/%Y %I:%M %p")
            + "'"
        )

        restricted = items.Restrict(restriction)

        result = {}
        for i in range(1, restricted.Count + 1):
            try:
                item = restricted.Item(i)
                if safe_get(item, "Class") != 26:  # olAppointment
                    continue

                # Mirrored standalone appointments carry the logical key of
                # their source occurrence. Native items derive their key from
                # Outlook recurrence identity. This is what makes both halves
                # of a pair appear under the same key during reconciliation.
                marker = property_accessor_get(item, SYNC_PROP)
                key = str(marker) if marker else occurrence_key(item)
                result[key] = item
            except Exception:
                logging.exception("Failed reading calendar item")
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

    def reconcile(self):
        a_view = self.expanded_items(self.calendar_a)
        b_view = self.expanded_items(self.calendar_b)
        rows = {r[0]: r for r in self.store.rows()}

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
                continue

            if ae is None:
                logging.info("A deleted %s; deleting B twin", key)
                self.delete_item(be)
                self.store.delete(key)
                continue

            if be is None:
                logging.info("B deleted %s; deleting A twin", key)
                self.delete_item(ae)
                self.store.delete(key)
                continue

            a_sig = signature(ae)
            b_sig = signature(be)
            a_changed = old_a_sig is not None and a_sig != old_a_sig
            b_changed = old_b_sig is not None and b_sig != old_b_sig

            if not a_changed and not b_changed:
                continue

            if a_changed and not b_changed:
                logging.info("Updating B from A: %s", key)
                self.apply_payload(ae, be)
                be.Save()
            elif b_changed and not a_changed:
                logging.info("Updating A from B: %s", key)
                self.apply_payload(be, ae)
                ae.Save()
            else:
                a_mod = modification_time(ae)
                b_mod = modification_time(be)
                if a_mod >= b_mod:
                    logging.warning("Conflict; A wins: %s", key)
                    self.apply_payload(ae, be)
                    be.Save()
                else:
                    logging.warning("Conflict; B wins: %s", key)
                    self.apply_payload(be, ae)
                    ae.Save()

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

        # New native A occurrences
        for key, ae in list(a_view.items()):
            if key in rows:
                continue
            if property_accessor_get(ae, SYNC_PROP):
                continue

            logging.info("Creating B twin: %s", key)
            be = self.copy_into_calendar(ae, self.calendar_b, key)

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

        # Re-read B because A->B creations were just added.
        b_view = self.expanded_items(self.calendar_b)

        # New native B occurrences
        for key, be in list(b_view.items()):
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
            try:
                self.reconcile()
            except Exception:
                logging.exception("Reconciliation failed")
            time.sleep(POLL_SECONDS)


def install_autostart():
    STARTUP_DIR.mkdir(parents=True, exist_ok=True)
    script_path = Path(__file__).resolve()
    python_exe = Path(sys.executable).resolve()

    content = (
        "@echo off\r\n"
        f'cd /d "{script_path.parent}"\r\n'
        f'start "" /min "{python_exe}" "{script_path}"\r\n'
    )
    STARTUP_CMD.write_text(content, encoding="utf-8")
    print(f"Installed per-user autostart: {STARTUP_CMD}")


def remove_autostart():
    if STARTUP_CMD.exists():
        STARTUP_CMD.unlink()
        print(f"Removed per-user autostart: {STARTUP_CMD}")
    else:
        print("Autostart entry was not present.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--install-autostart", action="store_true")
    parser.add_argument("--remove-autostart", action="store_true")
    args = parser.parse_args()

    if args.install_autostart:
        install_autostart()
        return

    if args.remove_autostart:
        remove_autostart()
        return

    setup_logging()
    sync = OutlookSync()
    sync.run()


if __name__ == "__main__":
    main()
