import unittest
from types import SimpleNamespace

import outlook_calendar_sync as syncmod


def test_settings(timeout=300, retry=10):
    return SimpleNamespace(
        STORE_A_HINT="alpha@example.com",
        STORE_B_HINT="bravo@example.com",
        OUTLOOK_READY_TIMEOUT_SECONDS=timeout,
        OUTLOOK_READY_RETRY_SECONDS=retry,
    )


class FakeClock:
    def __init__(self):
        self.now = 0
        self.sleeps = []

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


class OutlookStartupWaitTests(unittest.TestCase):
    def test_wait_retries_until_outlook_and_calendars_are_ready(self):
        clock = FakeClock()
        ready_sync = object()
        attempts = []

        def factory(settings, com_modules):
            attempts.append((settings, com_modules))
            if len(attempts) < 3:
                raise RuntimeError("stores are still loading")
            return ready_sync

        result = syncmod.wait_for_outlook_ready(
            test_settings(timeout=300, retry=10),
            com_modules=("pythoncom", "win32com.client"),
            sync_factory=factory,
            sleep_func=clock.sleep,
            monotonic_func=clock.monotonic,
        )

        self.assertIs(result, ready_sync)
        self.assertEqual(clock.sleeps, [10, 10])
        self.assertEqual(len(attempts), 3)

    def test_wait_stops_after_five_minutes_with_actionable_message(self):
        clock = FakeClock()
        attempts = []

        def factory(settings, com_modules):
            attempts.append(clock.now)
            raise RuntimeError("Could not find Outlook store matching 'alpha@example.com'")

        with self.assertRaises(syncmod.OutlookStartupTimeout) as caught:
            syncmod.wait_for_outlook_ready(
                test_settings(timeout=300, retry=60),
                com_modules=("pythoncom", "win32com.client"),
                sync_factory=factory,
                sleep_func=clock.sleep,
                monotonic_func=clock.monotonic,
            )

        message = str(caught.exception)
        self.assertIn("5 minutes", message)
        self.assertIn("alpha@example.com", message)
        self.assertIn("bravo@example.com", message)
        self.assertIn("Last error", message)
        self.assertEqual(clock.sleeps, [60, 60, 60, 60, 60])
        self.assertEqual(attempts, [0, 60, 120, 180, 240, 300])


if __name__ == "__main__":
    unittest.main()
