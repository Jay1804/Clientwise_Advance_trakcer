"""Local persistence for saved report configurations ("trackers").

Uses a SQLite file colocated with the app rather than the production MySQL
database (checkpoint_live) - that database is only ever queried (read) by
the rest of this app, and adding a table/write path there would be a much
bigger, harder-to-reverse change than this feature needs. Nothing in this
module ever touches the MySQL DB; it's a completely separate local file.

A tracker's date range is stored as `lookback_days` (a rolling window), not
frozen From/To dates - re-running a tracker later always pulls the last N
days as of that moment, rather than replaying a date range that ages out of
relevance (and eventually returns nothing).
"""

import getpass
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pandas as pd

DB_PATH = Path(__file__).parent / "trackers.db"

# Columns stored as JSON-encoded lists (or, for selected_columns, JSON null
# when no column customization has been saved yet).
_JSON_LIST_COLUMNS = (
    "client_ids",
    "client_labels",
    "check_names",
    "case_statuses",
    "check_statuses",
    "check_severities",
    "antecedent_fields",
    "schedule_days",
)
_JSON_COLUMNS = _JSON_LIST_COLUMNS + ("selected_columns",)

# Frequency values the UI/scheduler understand. "" means "no schedule - manual
# send only". schedule_days only applies to "Weekly"; the two "Daily (...)"
# options imply their own fixed day set.
SCHEDULE_FREQUENCIES = ("", "Daily (Mon-Fri)", "Daily (Mon-Sat)", "Weekly")
WEEKDAY_NAMES = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")

_ALL_COLUMNS = (
    "tracker_name",
    "tracker_description",
    "client_ids",
    "client_labels",
    "lookback_days",
    "check_names",
    "case_statuses",
    "check_statuses",
    "check_severities",
    "selected_columns",
    "antecedent_fields",
    "to_address",
    "cc_address",
    "schedule_frequency",
    "schedule_days",
    "schedule_time",
    "last_scheduled_run",
    "status",
    "created_at",
    "modified_at",
    "modified_by",
)

# Columns added after the table's first release - added via ALTER TABLE on an
# existing trackers.db rather than requiring a fresh DB, so trackers saved
# before this feature existed aren't lost.
_MIGRATION_COLUMNS = {
    "to_address": "TEXT NOT NULL DEFAULT ''",
    "cc_address": "TEXT NOT NULL DEFAULT ''",
    "schedule_frequency": "TEXT NOT NULL DEFAULT ''",
    "schedule_days": "TEXT NOT NULL DEFAULT '[]'",
    "schedule_time": "TEXT NOT NULL DEFAULT ''",
    "last_scheduled_run": "TEXT",
}


@contextmanager
def _connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS trackers (
                tracker_id INTEGER PRIMARY KEY AUTOINCREMENT,
                tracker_name TEXT NOT NULL,
                tracker_description TEXT NOT NULL DEFAULT '',
                client_ids TEXT NOT NULL DEFAULT '[]',
                client_labels TEXT NOT NULL DEFAULT '[]',
                lookback_days INTEGER NOT NULL DEFAULT 30,
                check_names TEXT NOT NULL DEFAULT '[]',
                case_statuses TEXT NOT NULL DEFAULT '[]',
                check_statuses TEXT NOT NULL DEFAULT '[]',
                check_severities TEXT NOT NULL DEFAULT '[]',
                selected_columns TEXT,
                antecedent_fields TEXT NOT NULL DEFAULT '[]',
                to_address TEXT NOT NULL DEFAULT '',
                cc_address TEXT NOT NULL DEFAULT '',
                schedule_frequency TEXT NOT NULL DEFAULT '',
                schedule_days TEXT NOT NULL DEFAULT '[]',
                schedule_time TEXT NOT NULL DEFAULT '',
                last_scheduled_run TEXT,
                status TEXT NOT NULL DEFAULT 'Active',
                created_at TEXT NOT NULL,
                modified_at TEXT NOT NULL,
                modified_by TEXT NOT NULL DEFAULT ''
            )
            """
        )
        existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(trackers)")}
        for col, decl in _MIGRATION_COLUMNS.items():
            if col not in existing_cols:
                conn.execute(f"ALTER TABLE trackers ADD COLUMN {col} {decl}")


def current_user() -> str:
    """OS login username, used as a stand-in for 'Modified By' - this app has
    no login system of its own."""
    try:
        return getpass.getuser()
    except Exception:
        return "unknown"


def _decode_row(row: sqlite3.Row) -> dict:
    d = dict(row)
    for col in _JSON_COLUMNS:
        raw = d.get(col)
        d[col] = json.loads(raw) if raw else ([] if col != "selected_columns" else None)
    return d


def save_tracker(config: dict, tracker_id: int | None = None) -> int:
    """Insert a new tracker, or update an existing one when `tracker_id` is
    given. Returns the tracker's id either way."""
    now = datetime.now().isoformat(timespec="seconds")
    row = {
        "tracker_name": config["tracker_name"],
        "tracker_description": config.get("tracker_description") or "",
        "lookback_days": int(config.get("lookback_days", 30)),
        "to_address": config.get("to_address") or "",
        "cc_address": config.get("cc_address") or "",
        "schedule_frequency": config.get("schedule_frequency") or "",
        "schedule_time": config.get("schedule_time") or "",
        "modified_at": now,
        "modified_by": current_user(),
    }
    for col in _JSON_LIST_COLUMNS:
        row[col] = json.dumps(config.get(col) or [])
    row["selected_columns"] = json.dumps(config.get("selected_columns"))

    with _connect() as conn:
        if tracker_id is None:
            row["created_at"] = now
            row["status"] = "Active"
            columns = ", ".join(row.keys())
            placeholders = ", ".join(f":{k}" for k in row.keys())
            cur = conn.execute(f"INSERT INTO trackers ({columns}) VALUES ({placeholders})", row)
            return cur.lastrowid
        else:
            set_clause = ", ".join(f"{k} = :{k}" for k in row.keys())
            row["tracker_id"] = tracker_id
            conn.execute(f"UPDATE trackers SET {set_clause} WHERE tracker_id = :tracker_id", row)
            return tracker_id


def get_tracker(tracker_id: int) -> dict | None:
    with _connect() as conn:
        row = conn.execute("SELECT * FROM trackers WHERE tracker_id = ?", (tracker_id,)).fetchone()
    return _decode_row(row) if row is not None else None


def set_status(tracker_id: int, status: str) -> None:
    with _connect() as conn:
        conn.execute(
            "UPDATE trackers SET status = ?, modified_at = ?, modified_by = ? WHERE tracker_id = ?",
            (status, datetime.now().isoformat(timespec="seconds"), current_user(), tracker_id),
        )


def update_recipients(tracker_id: int, to_address: str, cc_address: str) -> None:
    """Updates just the To/CC addresses, without needing the tracker's full
    config in hand - used by the inline 'Save recipients' action."""
    with _connect() as conn:
        conn.execute(
            "UPDATE trackers SET to_address = ?, cc_address = ?, modified_at = ?, modified_by = ? "
            "WHERE tracker_id = ?",
            (to_address, cc_address, datetime.now().isoformat(timespec="seconds"), current_user(), tracker_id),
        )


def mark_scheduled_run(tracker_id: int, when: datetime) -> None:
    """Records that a scheduled send happened, so the same slot isn't sent
    twice (see scheduled_tracker_runner.py). Deliberately doesn't bump
    modified_at/modified_by - an automatic send isn't a user edit."""
    with _connect() as conn:
        conn.execute(
            "UPDATE trackers SET last_scheduled_run = ? WHERE tracker_id = ?",
            (when.isoformat(timespec="seconds"), tracker_id),
        )


def list_trackers(
    name_query: str = "", client_query: str = "", status: str = "", client_id_query: str = ""
) -> pd.DataFrame:
    """All trackers (decoded), newest-modified first, optionally filtered.
    Filtering happens in Python after decoding since the interesting columns
    (client labels/ids) are JSON-encoded lists, not plain SQL-filterable text -
    fine at the scale a local single-user tool like this operates at."""
    init_db()
    with _connect() as conn:
        rows = conn.execute("SELECT * FROM trackers ORDER BY modified_at DESC").fetchall()

    records = [_decode_row(r) for r in rows]
    df = pd.DataFrame.from_records(records, columns=["tracker_id", *_ALL_COLUMNS]) if records else pd.DataFrame(
        columns=["tracker_id", *_ALL_COLUMNS]
    )
    if df.empty:
        return df

    if name_query:
        df = df[df["tracker_name"].str.contains(name_query, case=False, na=False)]
    if status:
        df = df[df["status"] == status]
    if client_query:
        needle = client_query.lower()
        df = df[df["client_labels"].apply(lambda labels: any(needle in str(l).lower() for l in labels))]
    if client_id_query:
        needle = client_id_query.lower()
        df = df[df["client_ids"].apply(lambda ids: any(needle in str(i).lower() for i in ids))]

    return df.reset_index(drop=True)
