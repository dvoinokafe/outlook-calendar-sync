"""
User settings for outlook_calendar_sync.py.

Edit this file for the two Outlook account/store names and sync timing.
Keep it next to outlook_calendar_sync.py.
"""

# Use distinctive text from each Classic Outlook store/mailbox display name.
# The email address is usually the easiest value to match.
# Email addresses also help skip duplicate copies of invitations sent
# between these two accounts.
STORE_A_HINT = "me@company-a.com"
STORE_B_HINT = "me@company-b.com"

# Calendar scan window.
DAYS_BACK = 2
DAYS_FORWARD = 2

# Responsiveness-friendly polling and scan pacing.
POLL_SECONDS = 5
SCAN_BATCH_SIZE = 25
SCAN_BATCH_PAUSE_SECONDS = 0.05

# Startup wait for Classic Outlook and both configured calendars.
OUTLOOK_READY_TIMEOUT_SECONDS = 300
OUTLOOK_READY_RETRY_SECONDS = 10
