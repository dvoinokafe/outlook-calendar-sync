import contextlib
import io
import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import outlook_calendar_sync as syncmod

ROOT = Path(__file__).resolve().parents[1]
MACRO_FILE = ROOT / "OutlookSyncControls.bas"


class FakePropertyAccessor:
    def __init__(self):
        self.values = {}

    def GetProperty(self, schema):
        if schema not in self.values:
            raise RuntimeError("property not found")
        return self.values[schema]

    def SetProperty(self, schema, value):
        self.values[schema] = value


class FakeItems:
    def __init__(self, folder):
        self.folder = folder

    def Add(self, _item_type):
        item = FakeAppointment(
            entry_id=f"{self.folder.StoreID}-new-{len(self.folder.created) + 1}",
            parent=self.folder,
            subject="",
        )
        self.folder.created.append(item)
        return item


class FakeFolder:
    def __init__(self, store_id, path):
        self.StoreID = store_id
        self.FolderPath = path
        self.created = []
        self.Items = FakeItems(self)


class FakeAppointment:
    def __init__(self, entry_id, parent, subject, marker=None):
        self.Class = 26
        self.EntryID = entry_id
        self.GlobalAppointmentID = entry_id
        self.Parent = parent
        self.Subject = subject
        self.Body = f"body {subject}"
        self.Start = datetime(2026, 1, 2, 9, 0)
        self.End = datetime(2026, 1, 2, 10, 0)
        self.AllDayEvent = False
        self.Location = "Office"
        self.Categories = ""
        self.Sensitivity = 0
        self.BusyStatus = 2
        self.Importance = 1
        self.ReminderSet = False
        self.ReminderMinutesBeforeStart = 15
        self.LastModificationTime = datetime(2026, 1, 2, 8, 0)
        self.PropertyAccessor = FakePropertyAccessor()
        self.save_count = 0
        if marker:
            self.PropertyAccessor.SetProperty(syncmod.SYNC_NS + syncmod.SYNC_PROP, marker)

    def Save(self):
        self.save_count += 1


class FakeStore:
    def __init__(self, rows=None):
        self._rows = rows or []
        self.upserts = []

    def rows(self):
        return list(self._rows)

    def upsert(self, *args):
        self.upserts.append(args)
        self._rows = [row for row in self._rows if row[0] != args[0]]
        self._rows.append(args)


class FakeNamespace:
    def __init__(self, items):
        self.items = items

    def GetItemFromID(self, entry_id, store_id):
        return self.items[(entry_id, store_id)]


def make_sync(rows=None, ns_items=None):
    sync = syncmod.OutlookSync.__new__(syncmod.OutlookSync)
    sync.calendar_a = FakeFolder("store-a", "\\A\\Calendar")
    sync.calendar_b = FakeFolder("store-b", "\\B\\Calendar")
    sync.store = FakeStore(rows=rows)
    sync.ns = FakeNamespace(ns_items or {})
    sync.settings = SimpleNamespace(POLL_SECONDS=5)
    return sync


class SelectedItemSyncTests(unittest.TestCase):
    def test_selected_a_item_creates_b_twin_when_pair_is_missing(self):
        sync = make_sync()
        selected = FakeAppointment("a-1", sync.calendar_a, "A source")

        result = sync.sync_selected_item(selected)

        created = sync.calendar_b.created[0]
        self.assertIn("Created copy in Calendar B", result)
        self.assertEqual(created.Subject, "A source")
        self.assertEqual(
            created.PropertyAccessor.GetProperty(syncmod.SYNC_NS + syncmod.SYNC_PROP),
            "single|a-1",
        )
        self.assertEqual(sync.store.upserts[0][1], "a-1")
        self.assertEqual(sync.store.upserts[0][3], created.EntryID)

    def test_selected_b_item_creates_a_twin_when_pair_is_missing(self):
        sync = make_sync()
        selected = FakeAppointment("b-1", sync.calendar_b, "B source")

        result = sync.sync_selected_item(selected)

        created = sync.calendar_a.created[0]
        self.assertIn("Created copy in Calendar A", result)
        self.assertEqual(created.Subject, "B source")
        self.assertEqual(sync.store.upserts[0][1], created.EntryID)
        self.assertEqual(sync.store.upserts[0][3], "b-1")

    def test_selected_item_updates_existing_opposite_twin(self):
        sync = make_sync(
            rows=[
                (
                    "single|a-2",
                    "a-2",
                    "store-a",
                    "b-2",
                    "store-b",
                    "old-a",
                    "old-b",
                    "old-a-mod",
                    "old-b-mod",
                )
            ]
        )
        selected = FakeAppointment("a-2", sync.calendar_a, "Updated A")
        twin = FakeAppointment("b-2", sync.calendar_b, "Old B")
        sync.ns = FakeNamespace({("b-2", "store-b"): twin})

        result = sync.sync_selected_item(selected)

        self.assertIn("Updated copy in Calendar B", result)
        self.assertEqual(twin.Subject, "Updated A")
        self.assertEqual(twin.save_count, 1)
        self.assertEqual(sync.store.upserts[-1][0], "single|a-2")

    def test_selected_mirrored_copy_updates_original_counterpart(self):
        key = "single|a-3"
        sync = make_sync(
            rows=[
                (
                    key,
                    "a-3",
                    "store-a",
                    "b-3",
                    "store-b",
                    "old-a",
                    "old-b",
                    "old-a-mod",
                    "old-b-mod",
                )
            ]
        )
        original = FakeAppointment("a-3", sync.calendar_a, "Old A")
        selected_copy = FakeAppointment("b-3", sync.calendar_b, "Edited mirror", marker=key)
        sync.ns = FakeNamespace({("a-3", "store-a"): original})

        result = sync.sync_selected_item(selected_copy)

        self.assertIn("Updated copy in Calendar A", result)
        self.assertEqual(original.Subject, "Edited mirror")
        self.assertEqual(original.save_count, 1)

    def test_selected_item_outside_configured_calendars_is_rejected(self):
        sync = make_sync()
        other_folder = FakeFolder("store-c", "\\C\\Calendar")
        selected = FakeAppointment("c-1", other_folder, "Outside")

        with self.assertRaises(syncmod.SelectedItemSyncError):
            sync.sync_selected_item(selected)

    def test_selected_item_in_other_folder_of_same_store_is_rejected(self):
        sync = make_sync()
        other_calendar = FakeFolder("store-a", "\\A\\Other Calendar")
        selected = FakeAppointment("a-other-1", other_calendar, "Wrong folder")

        with self.assertRaises(syncmod.SelectedItemSyncError):
            sync.sync_selected_item(selected)

    def test_selection_from_active_explorer_rejects_multiple_items(self):
        class Selection:
            Count = 2

            def Item(self, _index):
                return object()

        outlook = SimpleNamespace(
            ActiveInspector=lambda: None,
            ActiveExplorer=lambda: SimpleNamespace(Selection=Selection()),
        )

        with self.assertRaises(syncmod.SelectedItemSyncError):
            syncmod.selected_appointment_from_outlook(outlook)

    def test_main_sync_selected_runs_without_background_process(self):
        stdout = io.StringIO()

        with (
            mock.patch.object(sys, "argv", ["outlook_calendar_sync.py", "--sync-selected"]),
            mock.patch.object(
                syncmod,
                "sync_selected_item_once",
                return_value="Created copy in Calendar B.",
            ) as sync_once,
            contextlib.redirect_stdout(stdout),
        ):
            result = syncmod.main()

        self.assertEqual(result, 0)
        sync_once.assert_called_once_with(settings_path=None)
        self.assertIn("Created copy in Calendar B.", stdout.getvalue())

    def test_macro_exposes_sync_selected_button(self):
        text = MACRO_FILE.read_text(encoding="utf-8")

        self.assertIn("Public Sub OutlookSync_SyncSelected()", text)
        self.assertIn('MsgBox RunAndCapture("--sync-selected")', text)


if __name__ == "__main__":
    unittest.main()
