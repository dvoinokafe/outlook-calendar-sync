import unittest
from datetime import datetime
from types import SimpleNamespace

import outlook_calendar_sync as syncmod


class CountingBodyItem:
    def __init__(self, modified):
        self.Subject = "Planning"
        self.Start = datetime(2026, 1, 2, 9, 0, 0)
        self.End = datetime(2026, 1, 2, 9, 30, 0)
        self.LastModificationTime = modified
        self.body_reads = 0

    @property
    def Body(self):
        self.body_reads += 1
        return "Agenda"


class FakeStore:
    def __init__(self, rows):
        self._rows = list(rows)
        self.upserts = []
        self.deletes = []

    def rows(self):
        return list(self._rows)

    def upsert(self, *args):
        self.upserts.append(args)

    def delete(self, key):
        self.deletes.append(key)
        self._rows = [row for row in self._rows if row[0] != key]


class FakeCalendar:
    def __init__(self, name):
        self.name = name


class FakeAppointment:
    Class = 26
    GlobalAppointmentID = "gid-1"
    EntryID = "entry-1"
    RecurrenceState = syncmod.OL_RECURRENCE_STATE_NOT_RECURRING

    def __init__(self):
        self.Subject = "Planning"
        self.Start = datetime(2026, 1, 2, 9, 0, 0)
        self.End = datetime(2026, 1, 2, 9, 30, 0)
        self.LastModificationTime = datetime(2026, 1, 2, 8, 0, 0)


class CountMustNotBeUsedItems:
    def __init__(self, items):
        self._items = list(items)
        self._index = 0

    @property
    def Count(self):
        raise AssertionError("Outlook recurrence scans must not use Count")

    def Item(self, _index):
        raise AssertionError("Outlook recurrence scans must not use indexed Item")

    def GetFirst(self):
        self._index = 0
        if not self._items:
            return None
        return self._items[0]

    def GetNext(self):
        self._index += 1
        if self._index >= len(self._items):
            return None
        return self._items[self._index]


class FakeOutlookItems:
    def __init__(self, restricted_items):
        self.restricted_items = restricted_items
        self.IncludeRecurrences = False
        self.sort_key = None
        self.restriction = None

    def Sort(self, key):
        self.sort_key = key

    def Restrict(self, restriction):
        self.restriction = restriction
        return self.restricted_items


class FakeOutlookCalendar:
    def __init__(self, items):
        self.Items = items


class FakeSync:
    def __init__(self, a_view, b_view, rows):
        self.calendar_a = FakeCalendar("a")
        self.calendar_b = FakeCalendar("b")
        self.store = FakeStore(rows)
        self.views = {"a": a_view, "b": b_view}
        self.scan_counts = {"a": 0, "b": 0}

    def expanded_items(self, calendar):
        self.scan_counts[calendar.name] += 1
        return self.views[calendar.name]

    def delete_item(self, item):
        raise AssertionError("delete should not run in this scenario")

    def copy_into_calendar(self, source, target_calendar, logical_key):
        raise AssertionError("copy should not run in this scenario")


class DeleteTrackingSync(FakeSync):
    def __init__(self, a_view, b_view, rows):
        super().__init__(a_view, b_view, rows)
        self.deleted_items = []
        self.copied_keys = []

    def delete_item(self, item):
        self.deleted_items.append(item)

    def copy_into_calendar(self, source, target_calendar, logical_key):
        self.copied_keys.append(logical_key)
        raise AssertionError("stale deleted items must not be recreated")


class ResponsivenessTests(unittest.TestCase):
    def test_unchanged_pair_skips_expensive_body_reads(self):
        modified = datetime(2026, 1, 2, 3, 4, 5)
        stored_modified = "2026-01-02T03:04:05"
        key = "single|existing"
        a_item = CountingBodyItem(modified)
        b_item = CountingBodyItem(modified)
        sig = syncmod.signature(a_item)
        a_item.body_reads = 0
        b_item.body_reads = 0

        row = (
            key,
            "a-entry",
            "a-store",
            "b-entry",
            "b-store",
            sig,
            sig,
            stored_modified,
            stored_modified,
        )
        fake = FakeSync({key: a_item}, {key: b_item}, [row])

        syncmod.OutlookSync.reconcile(fake)

        self.assertEqual(fake.store.upserts, [])
        self.assertEqual(fake.store.deletes, [])
        self.assertEqual(a_item.body_reads, 0)
        self.assertEqual(b_item.body_reads, 0)

    def test_no_change_pass_scans_each_calendar_once(self):
        fake = FakeSync({}, {}, [])

        syncmod.OutlookSync.reconcile(fake)

        self.assertEqual(fake.scan_counts, {"a": 1, "b": 1})

    def test_expanded_items_iterates_recurring_view_without_count(self):
        appointment = FakeAppointment()
        restricted = CountMustNotBeUsedItems([appointment])
        items = FakeOutlookItems(restricted)
        calendar = FakeOutlookCalendar(items)
        fake_sync = SimpleNamespace(
            settings=SimpleNamespace(
                DAYS_BACK=2,
                DAYS_FORWARD=2,
                SCAN_BATCH_SIZE=0,
                SCAN_BATCH_PAUSE_SECONDS=0,
            )
        )

        result = syncmod.OutlookSync.expanded_items(fake_sync, calendar)

        self.assertEqual(result, {"single|gid-1": appointment})
        self.assertEqual(items.sort_key, "[Start]")
        self.assertTrue(items.IncludeRecurrences)
        self.assertIn("[Start] <=", items.restriction)

    def test_deleted_pair_is_not_recreated_from_stale_scan_view(self):
        modified = datetime(2026, 1, 2, 3, 4, 5)
        key = "single|stale-after-delete"
        a_item = CountingBodyItem(modified)
        row = (
            key,
            "a-entry",
            "a-store",
            "b-entry",
            "b-store",
            "a-sig",
            "b-sig",
            "2026-01-02T03:04:05",
            "2026-01-02T03:04:05",
        )
        fake = DeleteTrackingSync({key: a_item}, {}, [row])

        syncmod.OutlookSync.reconcile(fake)

        self.assertEqual(fake.store.deletes, [key])
        self.assertEqual(fake.deleted_items, [a_item])
        self.assertEqual(fake.copied_keys, [])


if __name__ == "__main__":
    unittest.main()
