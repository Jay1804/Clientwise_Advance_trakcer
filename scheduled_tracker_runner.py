"""Standalone runner: sends any Saved Tracker whose schedule is due right now.

This is a *checker*, not a daemon - it does one pass over trackers.db and
exits. Reliable "send this at 9am every weekday" behavior requires something
outside this app to invoke this script periodically (e.g. every 5-15 minutes)
- Windows Task Scheduler is the natural choice on the machine this runs on.
This script does not register that task itself; wire it up separately, e.g.:

    schtasks /Create /SC MINUTE /MO 10 /TN "AdvanceTracker Scheduled Sends" ^
        /TR "C:\\path\\to\\python.exe C:\\Gen AI\\AdvanceTracker_Client\\scheduled_tracker_runner.py" ^
        /RL LIMITED

Two important constraints on when this can actually work:
- It relies on Outlook COM automation (see outlook_email.py), which needs a
  logged-in Windows session with Outlook configured - it cannot run as a
  headless/locked-screen background service.
- A tracker is only ever sent once per scheduled day: `last_scheduled_run`
  is checked against "today" before sending, so re-running this script more
  often than the schedule's granularity is safe (it won't double-send), but
  the machine does need to be on, logged in, and this script does need to run
  at least once on/after the scheduled time for that day's send to happen.

Run manually to test: `python scheduled_tracker_runner.py`
"""

import sys
from datetime import datetime

import tracker_store as trackers
from outlook_email import outlook_available, send_via_outlook
from report_excel import build_report_excel
from report_runner import run_tracker_query


def _is_due(row, now: datetime) -> bool:
    freq = row["schedule_frequency"]
    if not freq or not row["schedule_time"] or row["status"] != "Active":
        return False

    today_name = trackers.WEEKDAY_NAMES[now.weekday()]
    if freq == "Daily (Mon-Fri)":
        due_today = today_name in trackers.WEEKDAY_NAMES[:5]
    elif freq == "Daily (Mon-Sat)":
        due_today = today_name in trackers.WEEKDAY_NAMES[:6]
    elif freq == "Weekly":
        due_today = today_name in (row["schedule_days"] or [])
    else:
        due_today = False
    if not due_today:
        return False

    try:
        sched_hour, sched_minute = (int(p) for p in row["schedule_time"].split(":"))
    except ValueError:
        return False
    scheduled_dt = now.replace(hour=sched_hour, minute=sched_minute, second=0, microsecond=0)

    last_run = row["last_scheduled_run"]
    if last_run:
        try:
            if datetime.fromisoformat(last_run).date() == now.date():
                return False  # already sent for today's slot
        except ValueError:
            pass

    # Due once we're at/after the scheduled time - catches this slot even if
    # the runner's polling interval means we're checking a few minutes late,
    # without ever firing early.
    return now >= scheduled_dt


def run_once() -> int:
    """Runs one pass. Returns the number of emails actually sent (for the
    caller/tests; the CLI entry point below just prints progress)."""
    if not outlook_available():
        print("Outlook automation (pywin32) isn't available in this environment - nothing to do.")
        return 0

    now = datetime.now()
    tracker_df = trackers.list_trackers(status="Active")
    sent = 0
    for _, row in tracker_df.iterrows():
        if not _is_due(row, now):
            continue

        tracker_id = int(row["tracker_id"])
        if not row["to_address"]:
            print(f"[skip] '{row['tracker_name']}' (#{tracker_id}) is due but has no To_address configured.")
            continue

        try:
            cfg = trackers.get_tracker(tracker_id)
            fresh_df, _, _ = run_tracker_query(cfg)
            file_bytes = build_report_excel(fresh_df, columns=cfg.get("selected_columns"))
            file_name = f"{row['tracker_name']}_{now:%Y%m%d}.xlsx"
            subject = f"Case Report - {row['tracker_name']} - {now:%Y-%m-%d}"
            body = f"Please find attached the scheduled '{row['tracker_name']}' report."
            send_via_outlook(row["to_address"], row["cc_address"], subject, body, file_bytes, file_name)
            trackers.mark_scheduled_run(tracker_id, now)
            sent += 1
            print(f"[sent] '{row['tracker_name']}' (#{tracker_id}) -> {row['to_address']}")
        except Exception as exc:
            print(f"[error] '{row['tracker_name']}' (#{tracker_id}): {exc}")

    if sent == 0:
        print(f"No trackers due at {now:%Y-%m-%d %H:%M}.")
    return sent


if __name__ == "__main__":
    sys.exit(0 if run_once() >= 0 else 1)
