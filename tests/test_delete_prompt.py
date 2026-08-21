import unittest
from datetime import datetime
from types import SimpleNamespace

import outlook_calendar_sync as syncmod


class FakePropertyAccessor:
    def __init__(self):
        self.values = {}

    def GetProperty(self, schema):
        if schema not in self.values:
            raise RuntimeError("property not found")
        return self.values[schema]

    def SetProperty(self, schema, value):
        self.values[schema] = value


class FakeStore:
    def __init__(self, rows):
        self._rows = list(rows)
        self.recorded_local_deletes = []

    def rows(self):
        return list(self._rows)

    def record_local_delete(self, logical_key, side):
        self.recorded_local_deletes.append((logical_key, side))


class FakeFolder:
    def __init__(self, store_id, path):
        self.StoreID = store_id
        self.EntryID = f"{store_id}:{path}"
        self.FolderPath = path


class FakeAppointment:
    Class = 26
    RecurrenceState = syncmod.OL_RECURRENCE_STATE_NOT_RECURRING

    def __init__(self, entry_id, parent, subject):
        self.EntryID = entry_id
        self.GlobalAppointmentID = entry_id
        self.Parent = parent
        self.Subject = subject
        self.Start = datetime(2026, 1, 2, 9, 0, 0)
        self.PropertyAccessor = FakePropertyAccessor()
        self.delete_calls = 0

    def Delete(self):
        self.delete_calls += 1


class DeletePromptSync(syncmod.OutlookSync):
    def __init__(self, prompt_result):
        self.calendar_a = FakeFolder("a-store", "\\Mailbox A\\Calendar")
        self.calendar_b = FakeFolder("b-store", "\\Mailbox B\\Calendar")
        self.store = FakeStore(
            [
                (
                    "single|a-entry",
                    "a-entry",
                    "a-store",
                    "b-entry",
                    "b-store",
                    "a-sig",
                    "b-sig",
                    "2026-01-02T03:04:05",
                    "2026-01-02T03:04:05",
                )
            ]
        )
        self.suppress_delete_prompt_depth = 0
        self.prompt_calls = []
        self.prompt_result = prompt_result

    def is_delete_destination(self, side, move_to):
        return move_to is None

    def prompt_delete_scope(self, item, side):
        self.prompt_calls.append((item.Subject, side))
        return self.prompt_result


class DeletePromptTests(unittest.TestCase):
    def test_local_only_choice_is_recorded_and_current_delete_continues(self):
        sync = DeletePromptSync(syncmod.DELETE_SCOPE_LOCAL)
        item = FakeAppointment("a-entry", sync.calendar_a, "Planning")

        cancel = sync.handle_before_item_move("a", item, None)

        self.assertFalse(cancel)
        self.assertEqual(sync.prompt_calls, [("Planning", "a")])
        self.assertEqual(sync.store.recorded_local_deletes, [("single|a-entry", "a")])

    def test_both_accounts_choice_allows_current_delete_without_local_marker(self):
        sync = DeletePromptSync(syncmod.DELETE_SCOPE_BOTH)
        item = FakeAppointment("a-entry", sync.calendar_a, "Planning")

        cancel = sync.handle_before_item_move("a", item, None)

        self.assertFalse(cancel)
        self.assertEqual(sync.prompt_calls, [("Planning", "a")])
        self.assertEqual(sync.store.recorded_local_deletes, [])

    def test_cancel_choice_blocks_current_delete(self):
        sync = DeletePromptSync(syncmod.DELETE_SCOPE_CANCEL)
        item = FakeAppointment("a-entry", sync.calendar_a, "Planning")

        cancel = sync.handle_before_item_move("a", item, None)

        self.assertTrue(cancel)
        self.assertEqual(sync.prompt_calls, [("Planning", "a")])
        self.assertEqual(sync.store.recorded_local_deletes, [])

    def test_script_initiated_delete_does_not_prompt(self):
        sync = DeletePromptSync(syncmod.DELETE_SCOPE_CANCEL)
        sync.suppress_delete_prompt_depth = 1
        item = FakeAppointment("a-entry", sync.calendar_a, "Planning")

        cancel = sync.handle_before_item_move("a", item, None)

        self.assertFalse(cancel)
        self.assertEqual(sync.prompt_calls, [])
        self.assertEqual(sync.store.recorded_local_deletes, [])

    def test_delete_item_suppresses_prompt_during_programmatic_delete(self):
        sync = DeletePromptSync(syncmod.DELETE_SCOPE_CANCEL)
        item = FakeAppointment("a-entry", sync.calendar_a, "Planning")
        depths = []

        def delete():
            depths.append(sync.suppress_delete_prompt_depth)
            item.delete_calls += 1

        item.Delete = delete

        sync.delete_item(item)

        self.assertEqual(item.delete_calls, 1)
        self.assertEqual(depths, [1])
        self.assertEqual(sync.suppress_delete_prompt_depth, 0)

    def test_unsynced_delete_does_not_prompt(self):
        sync = DeletePromptSync(syncmod.DELETE_SCOPE_CANCEL)
        item = FakeAppointment("unknown", sync.calendar_a, "Personal")

        cancel = sync.handle_before_item_move("a", item, None)

        self.assertFalse(cancel)
        self.assertEqual(sync.prompt_calls, [])
        self.assertEqual(sync.store.recorded_local_deletes, [])


class DeleteEventTests(unittest.TestCase):
    def test_folder_event_returns_cancel_result(self):
        sync = SimpleNamespace(handle_before_item_move=lambda side, item, move_to: True)
        event = syncmod.OutlookFolderDeleteEvents()
        event.sync = sync
        event.side = "b"

        result = event.OnBeforeItemMove(object(), None, False)

        self.assertTrue(result)


if __name__ == "__main__":
    unittest.main()
