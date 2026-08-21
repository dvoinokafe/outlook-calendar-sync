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


class FakeRecipient:
    def __init__(self, address):
        self.Address = address
        self.Name = address


class FakeRecipients:
    def __init__(self, addresses):
        self._items = [FakeRecipient(address) for address in addresses]

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class FakeAppointment:
    def __init__(
        self,
        entry_id,
        parent,
        subject,
        marker=None,
        organizer="",
        recipients=None,
        required_attendees="",
        optional_attendees="",
    ):
        self.Class = 26
        self.EntryID = entry_id
        self.GlobalAppointmentID = entry_id
        self.Parent = parent
        self.Subject = subject
        self.Organizer = organizer
        self.RequiredAttendees = required_attendees
        self.OptionalAttendees = optional_attendees
        self.Recipients = FakeRecipients(recipients or [])
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


class FakeCategories:
    def __init__(self, existing=None):
        self.existing = set(existing or [])
        self.added = []

    def Item(self, name):
        if name not in self.existing:
            raise RuntimeError("category not found")
        return SimpleNamespace(Name=name)

    def Add(self, name, color, shortcut):
        self.existing.add(name)
        self.added.append((name, color, shortcut))
        return SimpleNamespace(Name=name, Color=color, ShortcutKey=shortcut)


class FakeNamespace:
    def __init__(self, items, categories=None):
        self.items = items
        self.Categories = categories or FakeCategories()

    def GetItemFromID(self, entry_id, store_id):
        return self.items[(entry_id, store_id)]


def make_sync(rows=None, ns_items=None):
    sync = syncmod.OutlookSync.__new__(syncmod.OutlookSync)
    sync.calendar_a = FakeFolder("store-a", "\\A\\Calendar")
    sync.calendar_b = FakeFolder("store-b", "\\B\\Calendar")
    sync.store = FakeStore(rows=rows)
    sync.ns = FakeNamespace(ns_items or {})
    sync.settings = SimpleNamespace(
        POLL_SECONDS=5,
        STORE_A_HINT="account-a@example.com",
        STORE_B_HINT="account-b@example.com",
    )
    sync.settings_path = Path("C:/Config/settings.py")
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
        self.assertEqual(created.Categories, "DY Sync")
        self.assertEqual(sync.store.upserts[0][1], "a-1")
        self.assertEqual(sync.store.upserts[0][3], created.EntryID)

    def test_selected_b_item_creates_a_twin_when_pair_is_missing(self):
        sync = make_sync()
        selected = FakeAppointment("b-1", sync.calendar_b, "B source")

        result = sync.sync_selected_item(selected)

        created = sync.calendar_a.created[0]
        self.assertIn("Created copy in Calendar A", result)
        self.assertEqual(created.Subject, "B source")
        self.assertEqual(created.Categories, "DY Sync")
        self.assertEqual(sync.store.upserts[0][1], created.EntryID)
        self.assertEqual(sync.store.upserts[0][3], "b-1")

    def test_selected_a_invite_to_b_is_skipped(self):
        sync = make_sync()
        selected = FakeAppointment(
            "a-invite",
            sync.calendar_a,
            "A invites B",
            organizer="account-a@example.com",
            recipients=["account-b@example.com"],
        )

        result = sync.sync_selected_item(selected)

        self.assertIn("Skipped cross-account invite", result)
        self.assertEqual(sync.calendar_b.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_selected_b_invite_to_a_is_skipped(self):
        sync = make_sync()
        selected = FakeAppointment(
            "b-invite",
            sync.calendar_b,
            "B invites A",
            organizer="account-b@example.com",
            recipients=["account-a@example.com"],
        )

        result = sync.sync_selected_item(selected)

        self.assertIn("Skipped cross-account invite", result)
        self.assertEqual(sync.calendar_a.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_selected_forwarded_a_invite_to_b_is_skipped(self):
        sync = make_sync()
        selected = FakeAppointment(
            "a-forwarded-invite",
            sync.calendar_a,
            "A forwards external invite to B",
            organizer="external@example.com",
            recipients=["account-b@example.com"],
        )

        result = sync.sync_selected_item(selected)

        self.assertIn("Skipped cross-account invite", result)
        self.assertEqual(sync.calendar_b.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_selected_forwarded_b_invite_to_a_is_skipped(self):
        sync = make_sync()
        selected = FakeAppointment(
            "b-forwarded-invite",
            sync.calendar_b,
            "B forwards external invite to A",
            organizer="external@example.com",
            recipients=["account-a@example.com"],
        )

        result = sync.sync_selected_item(selected)

        self.assertIn("Skipped cross-account invite", result)
        self.assertEqual(sync.calendar_a.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_selected_unrelated_meeting_still_creates_copy(self):
        sync = make_sync()
        selected = FakeAppointment(
            "a-external",
            sync.calendar_a,
            "A invites external",
            organizer="account-a@example.com",
            recipients=["external@example.com"],
        )

        result = sync.sync_selected_item(selected)

        self.assertIn("Created copy in Calendar B", result)
        created = sync.calendar_b.created[0]
        self.assertEqual(created.Categories, "DY Sync")
        self.assertEqual(len(sync.calendar_b.created), 1)
        self.assertEqual(len(sync.store.upserts), 1)

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
        twin = FakeAppointment("b-2", sync.calendar_b, "Old B", marker="single|a-2")
        sync.ns = FakeNamespace({("b-2", "store-b"): twin})

        result = sync.sync_selected_item(selected)

        self.assertIn("Updated copy in Calendar B", result)
        self.assertEqual(twin.Subject, "Updated A")
        self.assertEqual(twin.Categories, "DY Sync")
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
        selected_copy.Categories = "Work, DY Sync"
        sync.ns = FakeNamespace({("a-3", "store-a"): original})

        result = sync.sync_selected_item(selected_copy)

        self.assertIn("Updated copy in Calendar A", result)
        self.assertEqual(original.Subject, "Edited mirror")
        self.assertEqual(original.Categories, "Work")
        self.assertEqual(original.save_count, 1)

    def test_selected_mirrored_copy_uses_marker_when_folder_parent_is_unexpected(self):
        key = "single|a-4"
        sync = make_sync(
            rows=[
                (
                    key,
                    "a-4",
                    "store-a",
                    "b-4",
                    "store-b",
                    "old-a",
                    "old-b",
                    "old-a-mod",
                    "old-b-mod",
                )
            ]
        )
        original = FakeAppointment("a-4", sync.calendar_a, "Old A")
        search_folder = FakeFolder("search-store", "\\Search Results")
        selected_copy = FakeAppointment(
            "b-4",
            search_folder,
            "Edited mirror from search",
            marker=key,
        )
        sync.ns = FakeNamespace({("a-4", "store-a"): original})

        result = sync.sync_selected_item(selected_copy)

        self.assertIn("Updated copy in Calendar A", result)
        self.assertEqual(original.Subject, "Edited mirror from search")
        self.assertEqual(original.save_count, 1)

    def test_background_sync_skips_a_invite_to_b(self):
        sync = make_sync()
        item = FakeAppointment(
            "a-background-invite",
            sync.calendar_a,
            "A background invite",
            organizer="account-a@example.com",
            recipients=["account-b@example.com"],
        )
        sync.expanded_items = lambda calendar: (
            {"single|a-background-invite": item}
            if calendar is sync.calendar_a
            else {}
        )

        sync.reconcile()

        self.assertEqual(sync.calendar_b.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_background_sync_skips_forwarded_a_invite_to_b(self):
        sync = make_sync()
        item = FakeAppointment(
            "a-background-forwarded-invite",
            sync.calendar_a,
            "A background forwarded invite",
            organizer="external@example.com",
            recipients=["account-b@example.com"],
        )
        sync.expanded_items = lambda calendar: (
            {"single|a-background-forwarded-invite": item}
            if calendar is sync.calendar_a
            else {}
        )

        sync.reconcile()

        self.assertEqual(sync.calendar_b.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_background_sync_skips_b_invite_to_a(self):
        sync = make_sync()
        item = FakeAppointment(
            "b-background-invite",
            sync.calendar_b,
            "B background invite",
            organizer="account-b@example.com",
            recipients=["account-a@example.com"],
        )
        sync.expanded_items = lambda calendar: (
            {"single|b-background-invite": item}
            if calendar is sync.calendar_b
            else {}
        )

        sync.reconcile()

        self.assertEqual(sync.calendar_a.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_background_sync_skips_forwarded_b_invite_to_a(self):
        sync = make_sync()
        item = FakeAppointment(
            "b-background-forwarded-invite",
            sync.calendar_b,
            "B background forwarded invite",
            organizer="external@example.com",
            recipients=["account-a@example.com"],
        )
        sync.expanded_items = lambda calendar: (
            {"single|b-background-forwarded-invite": item}
            if calendar is sync.calendar_b
            else {}
        )

        sync.reconcile()

        self.assertEqual(sync.calendar_a.created, [])
        self.assertEqual(sync.store.upserts, [])

    def test_background_sync_still_copies_unrelated_meeting(self):
        sync = make_sync()
        item = FakeAppointment(
            "a-background-external",
            sync.calendar_a,
            "A external invite",
            organizer="account-a@example.com",
            recipients=["external@example.com"],
        )
        sync.expanded_items = lambda calendar: (
            {"single|a-background-external": item}
            if calendar is sync.calendar_a
            else {}
        )

        sync.reconcile()

        created = sync.calendar_b.created[0]
        self.assertEqual(created.Categories, "DY Sync")
        self.assertEqual(len(sync.calendar_b.created), 1)
        self.assertEqual(len(sync.store.upserts), 1)

    def test_signature_ignores_sync_category_marker(self):
        plain = FakeAppointment("sig-1", FakeFolder("store-a", "\\A\\Calendar"), "Same")
        marked = FakeAppointment("sig-1", FakeFolder("store-a", "\\A\\Calendar"), "Same")
        plain.Categories = "Work"
        marked.Categories = "Work, DY Sync"

        self.assertEqual(syncmod.signature(plain), syncmod.signature(marked))

    def test_ensure_sync_category_creates_gray_outlook_category(self):
        categories = FakeCategories()
        namespace = FakeNamespace({}, categories=categories)

        getattr(syncmod, "ensure_sync_category", lambda _namespace: None)(namespace)

        self.assertEqual(categories.added, [("DY Sync", 13, 0)])

    def test_ensure_sync_category_reuses_existing_category(self):
        categories = FakeCategories(existing=["DY Sync"])
        namespace = FakeNamespace({}, categories=categories)

        getattr(syncmod, "ensure_sync_category", lambda _namespace: None)(namespace)

        self.assertEqual(categories.added, [])

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

    def test_selected_item_rejection_includes_folder_diagnostics(self):
        sync = make_sync()
        other_folder = FakeFolder("store-c", "\\C\\Calendar")
        selected = FakeAppointment("c-2", other_folder, "Outside")

        with self.assertRaises(syncmod.SelectedItemSyncError) as caught:
            sync.sync_selected_item(selected)

        message = str(caught.exception)
        self.assertIn("Selected subject: Outside", message)
        self.assertIn("Selected folder: \\C\\Calendar", message)
        self.assertIn("Selected store: store-c", message)
        self.assertIn("Calendar A: \\A\\Calendar", message)
        self.assertIn("Calendar B: \\B\\Calendar", message)
        self.assertIn("Settings: C:\\Config\\settings.py", message)

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
