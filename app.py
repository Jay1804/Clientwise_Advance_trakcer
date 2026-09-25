"""Streamlit app: Client Wise Case/Check report, downloadable consolidated or split by client."""

from datetime import date, datetime, time, timedelta

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from streamlit_sortables import sort_items

from db import run_query_params
from flex_fields import load_flex_field_map
from queries import (
    CASE_LEVEL_COLUMNS,
    CASE_REPORT_QUERY,
    CASE_STATUS_FILTER_CLAUSE,
    CASE_STATUS_LOOKUP_QUERY,
    CHECK_LEVEL_COLUMNS,
    CHECK_NAME_FILTER_CLAUSE,
    CHECK_NAME_LOOKUP_BY_CLIENT_QUERY,
    CHECK_NAME_LOOKUP_QUERY,
    CHECK_SEVERITY_FILTER_CLAUSE,
    CHECK_SEVERITY_LOOKUP_BY_CLIENT_QUERY,
    CHECK_SEVERITY_LOOKUP_QUERY,
    CHECK_STATUS_FILTER_CLAUSE,
    CHECK_STATUS_LOOKUP_QUERY,
    CLIENT_FILTER_CLAUSE,
    CLIENT_LOOKUP_QUERY,
    FIELD_NAME_LOOKUP_QUERY,
)
from outlook_email import outlook_available, send_via_outlook
from report_excel import (
    MAX_SPLIT_GROUPS,
    TooManySplitGroupsError,
    aggregate_selected_columns,
    build_group_split_zip,
    build_mapping_template_excel,
    build_report_excel,
    build_single_group_excel,
    build_split_report_zip,
    sanitize_filename,
)
from report_runner import fetch_antecedent_columns, run_tracker_query
import tracker_store as trackers

load_dotenv()
trackers.init_db()

st.set_page_config(page_title="Advance Tracker - Client Report", layout="wide")
st.title("Advance Tracker - Client Report")

BUILDER_VIEW = "📝 Report Builder"
TRACKERS_VIEW = "💾 Saved Trackers"

LARGE_RANGE_NO_FILTER_DAYS = 31


@st.cache_data(ttl=3600)
def get_client_lookup() -> pd.DataFrame:
    return run_query_params(CLIENT_LOOKUP_QUERY)


@st.cache_data(ttl=3600)
def get_check_name_lookup(client_ids: tuple[int, ...] = ()) -> pd.DataFrame:
    """Unique check names. With no client selected, this reads the small
    master check catalog directly; with clients selected, it's scoped to
    check names actually used on those clients' cases (a heavier join)."""
    if client_ids:
        return run_query_params(CHECK_NAME_LOOKUP_BY_CLIENT_QUERY, {"client_ids": list(client_ids)})
    return run_query_params(CHECK_NAME_LOOKUP_QUERY)


@st.cache_data(ttl=3600)
def get_case_status_lookup() -> pd.DataFrame:
    return run_query_params(CASE_STATUS_LOOKUP_QUERY)


@st.cache_data(ttl=3600)
def get_check_status_lookup() -> pd.DataFrame:
    return run_query_params(CHECK_STATUS_LOOKUP_QUERY)


@st.cache_data(ttl=3600)
def get_check_severity_lookup(client_ids: tuple[int, ...] = ()) -> pd.DataFrame:
    """Distinct check severities. Scoped to selected clients this is a quick
    query (~2-3s); with no client selected it's a 50s+ full-table scan (see
    queries.py), so that path is only ever run on explicit request."""
    if client_ids:
        return run_query_params(CHECK_SEVERITY_LOOKUP_BY_CLIENT_QUERY, {"client_ids": list(client_ids)})
    return run_query_params(CHECK_SEVERITY_LOOKUP_QUERY)


@st.cache_data(ttl=3600)
def get_field_name_lookup(case_check_ids: tuple[int, ...]) -> pd.DataFrame:
    """Antecedent field names actually used by these specific checks - scoped
    to the report's own Case_Check_id values (not by client) since that skips
    the client join chain entirely and stays fast (~3-6s) regardless of how
    many clients the report spans (see queries.py)."""
    if not case_check_ids:
        return pd.DataFrame({"field_name": []})
    return run_query_params(FIELD_NAME_LOOKUP_QUERY, {"case_check_ids": list(case_check_ids)})


if "check_severity_options" not in st.session_state:
    st.session_state.check_severity_options = None
if "distinct_client_ids" not in st.session_state:
    st.session_state.distinct_client_ids = []
if "enable_antecedents" not in st.session_state:
    st.session_state.enable_antecedents = False
if "antecedent_columns" not in st.session_state:
    st.session_state.antecedent_columns = []
if "antecedent_field_types" not in st.session_state:
    st.session_state.antecedent_field_types = {}

GROUP_SPLIT_SESSION_KEYS = (
    "split_by_process_zip",
    "split_by_location_zip",
    "split_by_flexfield_zip",
    "split_by_antecedent_zip",
)
for _key in GROUP_SPLIT_SESSION_KEYS:
    if _key not in st.session_state:
        st.session_state[_key] = None

if "view_switch" not in st.session_state:
    st.session_state.view_switch = BUILDER_VIEW
if "pending_view" not in st.session_state:
    st.session_state.pending_view = None
# Streamlit refuses to let you set session_state[key] for a widget that has
# already been instantiated in this same run - so Edit/Email Tracker (called
# from render_saved_trackers_view(), i.e. *after* the segmented_control below
# has already rendered) can't just assign st.session_state.view_switch
# directly. They stash the target view here instead, applied on the next
# run's very first line, strictly before that widget renders again.
if st.session_state.pending_view is not None:
    st.session_state.view_switch = st.session_state.pending_view
    st.session_state.pending_view = None
# Seeded here (rather than via `value=` on the widgets below) so a tracker
# load can freely overwrite them beforehand without Streamlit's "widget had
# both a default value and a Session State value" warning.
if "flt_from_date" not in st.session_state:
    st.session_state.flt_from_date = date.today() - timedelta(days=30)
if "flt_to_date" not in st.session_state:
    st.session_state.flt_to_date = date.today()
if "loaded_tracker_id" not in st.session_state:
    st.session_state.loaded_tracker_id = None
if "loaded_tracker_columns" not in st.session_state:
    st.session_state.loaded_tracker_columns = None
if "loaded_tracker_name" not in st.session_state:
    st.session_state.loaded_tracker_name = ""
if "loaded_tracker_description" not in st.session_state:
    st.session_state.loaded_tracker_description = ""
if "loaded_tracker_to_address" not in st.session_state:
    st.session_state.loaded_tracker_to_address = ""
if "loaded_tracker_cc_address" not in st.session_state:
    st.session_state.loaded_tracker_cc_address = ""
if "loaded_tracker_schedule_enabled" not in st.session_state:
    st.session_state.loaded_tracker_schedule_enabled = False
if "loaded_tracker_schedule_frequency" not in st.session_state:
    st.session_state.loaded_tracker_schedule_frequency = ""
if "loaded_tracker_schedule_days" not in st.session_state:
    st.session_state.loaded_tracker_schedule_days = []
if "loaded_tracker_schedule_time" not in st.session_state:
    st.session_state.loaded_tracker_schedule_time = time(9, 0)
if "pending_tracker_load" not in st.session_state:
    st.session_state.pending_tracker_load = None
if "tracker_page" not in st.session_state:
    st.session_state.tracker_page = 0

if "report_df" not in st.session_state:
    st.session_state.report_df = None
if "report_excel_bytes" not in st.session_state:
    st.session_state.report_excel_bytes = None
if "report_zip_bytes" not in st.session_state:
    st.session_state.report_zip_bytes = None
if "flex_map" not in st.session_state:
    st.session_state.flex_map = {}
if "selected_columns" not in st.session_state:
    st.session_state.selected_columns = None
if "available_columns" not in st.session_state:
    st.session_state.available_columns = None
if "applied_columns" not in st.session_state:
    st.session_state.applied_columns = None
if "sorter_version" not in st.session_state:
    st.session_state.sorter_version = 0


def _merge_column_lists(columns: list[str]) -> tuple[list[str], list[str]]:
    """(selected, available) for the column picker, preserving the user's
    existing choices/order for columns still present and appending any new
    ones as selected (kept simple since this app's query has a fixed column
    list - this just guards against it changing under us)."""
    prev_selected = st.session_state.selected_columns
    if prev_selected is None:
        return list(columns), []

    prev_available = st.session_state.available_columns or []
    known = set(prev_selected) | set(prev_available)
    kept_selected = [c for c in prev_selected if c in columns]
    kept_available = [c for c in prev_available if c in columns]
    new_cols = [c for c in columns if c not in known]
    return kept_selected + new_cols, kept_available


def _rebuild_exports(df: pd.DataFrame, columns: list[str], flex_map: dict) -> tuple[bytes, bytes]:
    return build_report_excel(df, columns=columns), build_split_report_zip(df, flex_map, columns=columns)


def _parse_mapping_file(uploaded_file, dimension_label: str) -> pd.DataFrame:
    """Read an uploaded mapping file (.xlsx or .csv) and return a DataFrame
    with exactly [dimension_label, To_address, CC_address] - trimmed, with
    blank-dimension rows dropped. Raises ValueError with a user-facing
    message if a required column is missing."""
    name = (uploaded_file.name or "").lower()
    mapping_df = pd.read_csv(uploaded_file, dtype=str) if name.endswith(".csv") else pd.read_excel(uploaded_file, dtype=str)
    mapping_df.columns = [str(c).strip() for c in mapping_df.columns]

    required = {dimension_label, "To_address", "CC_address"}
    missing = required - set(mapping_df.columns)
    if missing:
        raise ValueError(f"missing column(s): {', '.join(sorted(missing))}")

    mapping_df = mapping_df[[dimension_label, "To_address", "CC_address"]].copy()
    for col in mapping_df.columns:
        mapping_df[col] = mapping_df[col].fillna("").astype(str).str.strip()
    return mapping_df[mapping_df[dimension_label] != ""]


def _render_group_split_download(
    label: str,
    dimension_label: str,
    group_column: str,
    session_key: str,
    df: pd.DataFrame,
    columns: list[str],
) -> None:
    """One 'Generate' button that builds a split-by-`group_column` zip into
    `session_key`, a download button once it's built, and an email-via-Outlook
    sub-section keyed on a user-uploaded To/CC mapping file. Split on demand
    rather than automatically since some columns (e.g. a free-text antecedent
    field) can have far too many distinct values to split by.

    `dimension_label` is the human name used in the mapping file's header and
    prompts (e.g. "Case Flex Field") - separate from `label`, which can carry
    the specific chosen column too (e.g. "Flex Field (CASE_FLEX_FIELD1)")."""
    col_generate, col_download = st.columns(2)
    with col_generate:
        if st.button(f"🔀 Generate split by {label}", key=f"gen_{session_key}", use_container_width=True):
            try:
                with st.spinner(f"Building split-by-{label} report..."):
                    st.session_state[session_key] = build_group_split_zip(df, group_column, columns=columns)
            except TooManySplitGroupsError as exc:
                st.session_state[session_key] = None
                st.error(str(exc))
    with col_download:
        if st.session_state.get(session_key):
            st.download_button(
                label=f"⬇️ Download Split by {label} (.zip)",
                data=st.session_state[session_key],
                file_name=f"Client_Case_Report_Split_By_{label.replace(' ', '_')}_{date.today().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                use_container_width=True,
                key=f"dl_{session_key}",
            )

    with st.expander(f"📧 Email split-by-{label} files via Outlook", expanded=True):
        if group_column not in df.columns:
            st.caption("This column isn't in the current report.")
            return
        distinct_values = sorted(df[group_column].dropna().unique().tolist(), key=str)
        if not distinct_values:
            st.caption(
                f"'{label}' has no non-empty values in this report, so there's "
                "nothing to split/email by - pick a different one above."
            )
            return

        if not outlook_available():
            st.warning(
                "Outlook automation (pywin32) isn't available in this environment, "
                "so emailing is disabled here."
            )
            return

        template_bytes = build_mapping_template_excel(dimension_label, distinct_values)
        st.download_button(
            label=f"📄 Download {dimension_label} mapping template (.xlsx)",
            data=template_bytes,
            file_name=f"Mapping_Template_{dimension_label.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{session_key}_template_dl",
        )
        st.caption(
            f"Prefilled with every {dimension_label} value currently in this report. Fill in "
            "'To_address' for the rows you want emailed (required) and 'CC_address' "
            "(optional; separate multiple addresses with a semicolon), then upload it below."
        )

        uploaded = st.file_uploader(
            f"Upload completed {dimension_label} mapping file (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key=f"{session_key}_mapping_upload",
        )
        if uploaded is None:
            return

        try:
            mapping_df = _parse_mapping_file(uploaded, dimension_label)
        except Exception as exc:
            st.error(f"Couldn't read mapping file - {exc}")
            return

        sendable = mapping_df[mapping_df["To_address"] != ""]
        known_values = {str(v) for v in distinct_values}
        unmatched = sendable[~sendable[dimension_label].isin(known_values)]
        st.caption(
            f"{len(sendable)} row(s) with a To_address."
            + (
                f" {len(unmatched)} of them don't match any {dimension_label} value in "
                "this report and will be skipped."
                if len(unmatched)
                else ""
            )
        )

        subject_template = st.text_input(
            "Email subject", value=f"Case Report - {dimension_label}: {{value}}", key=f"{session_key}_subject"
        )
        body_template = st.text_area(
            "Email body",
            value=f"Please find attached the case report for {dimension_label} = {{value}}.",
            key=f"{session_key}_body",
        )

        if st.button(
            f"📧 Send {len(sendable)} email(s) via Outlook",
            key=f"{session_key}_send",
            disabled=sendable.empty,
        ):
            value_by_str = {str(v): v for v in distinct_values}
            results = []
            progress = st.progress(0.0)
            for i, (_, row) in enumerate(sendable.iterrows()):
                display_value = row[dimension_label]
                progress.progress((i + 1) / len(sendable))
                if display_value not in value_by_str:
                    results.append({dimension_label: display_value, "To_address": row["To_address"], "Status": "Skipped - no matching data"})
                    continue
                try:
                    actual_value = value_by_str[display_value]
                    attachment_bytes = build_single_group_excel(df, group_column, actual_value, columns=columns)
                    file_name = sanitize_filename(display_value, set(), max_len=150, fallback="Report") + ".xlsx"
                    subject = subject_template.replace("{value}", display_value)
                    body = body_template.replace("{value}", display_value)
                    send_via_outlook(row["To_address"], row["CC_address"], subject, body, attachment_bytes, file_name)
                    results.append({dimension_label: display_value, "To_address": row["To_address"], "Status": "Sent"})
                except Exception as exc:
                    results.append({dimension_label: display_value, "To_address": row["To_address"], "Status": f"Failed - {exc}"})
            st.dataframe(pd.DataFrame(results), use_container_width=True)


TRACKER_PAGE_SIZE = 10


def render_saved_trackers_view() -> None:
    """The 'Saved Trackers' view: search/filter saved report configs and
    Edit/Email/Download/Activate/Deactivate them, without re-entering all the
    filters from scratch each time."""
    st.subheader("Saved Trackers")

    with st.form("tracker_search_form", border=True):
        col_name, col_client, col_status = st.columns(3)
        with col_name:
            st.text_input("Tracker Name", key="ts_name_query", placeholder="Search by name")
        with col_client:
            st.text_input("Client", key="ts_client_query", placeholder="Search by client name")
        with col_status:
            st.selectbox(
                "Status",
                options=["", "Active", "Inactive"],
                format_func=lambda v: v or "-Filter by Status-",
                key="ts_status_query",
            )
        st.text_input("Client External Code", key="ts_client_id_query", placeholder="Search by Client External Code")
        st.form_submit_button("🔍 Search")

    tracker_df = trackers.list_trackers(
        name_query=st.session_state.get("ts_name_query", ""),
        client_query=st.session_state.get("ts_client_query", ""),
        status=st.session_state.get("ts_status_query", ""),
        client_id_query=st.session_state.get("ts_client_id_query", ""),
    )

    if tracker_df.empty:
        st.info(
            "No saved trackers match this search. Build a report on the "
            f"'{BUILDER_VIEW}' view and save it from there - or clear these filters."
        )
        return

    total_pages = max(1, -(-len(tracker_df) // TRACKER_PAGE_SIZE))
    st.session_state.tracker_page = max(0, min(st.session_state.tracker_page, total_pages - 1))
    start = st.session_state.tracker_page * TRACKER_PAGE_SIZE
    page_df = tracker_df.iloc[start : start + TRACKER_PAGE_SIZE]

    info_widths = [0.4, 1, 1.8, 1.6, 1, 1.3, 1]
    for col, title in zip(
        st.columns(info_widths),
        ["S.No", "Client Ext Code", "Client", "Tracker Name", "Created", "Modified", "Modified By"],
    ):
        col.markdown(f"**{title}**")
    st.divider()

    for i, (_, row) in enumerate(page_df.iterrows()):
        tracker_id = int(row["tracker_id"])
        with st.container(border=True):
            info_cols = st.columns(info_widths)
            info_cols[0].write(start + i + 1)
            info_cols[1].write(", ".join(str(c) for c in row["client_ids"]) or "All")
            info_cols[2].write(", ".join(row["client_labels"]) or "All clients")
            info_cols[3].markdown(f"**{row['tracker_name']}** ({row['status']})")
            info_cols[4].write(str(row["created_at"])[:10])
            info_cols[5].write(str(row["modified_at"])[:16].replace("T", " "))
            info_cols[6].write(row["modified_by"])

            if row["tracker_description"]:
                st.caption(row["tracker_description"])

            if row["schedule_frequency"]:
                if row["schedule_frequency"] == "Weekly":
                    sched_desc = f"Weekly on {', '.join(row['schedule_days']) or 'no days set'}"
                else:
                    sched_desc = row["schedule_frequency"]
                last_run = row["last_scheduled_run"]
                last_run_label = f", last sent {str(last_run)[:16].replace('T', ' ')}" if last_run else ", never sent yet"
                st.caption(f"🕒 Scheduled: {sched_desc} at {row['schedule_time'] or '?'}{last_run_label}")
            else:
                st.caption("🕒 Not scheduled - manual send only")

            # Full-width columns here (not nested inside a narrow parent
            # column) so each button gets enough room for its label - nesting
            # 5 buttons inside one already-narrow "Action" column squeezed
            # each one down to a couple of characters wide.
            action_cols = st.columns(6)
            if action_cols[0].button("✏️ Edit", key=f"tracker_edit_{tracker_id}", help="Load into Report Builder", use_container_width=True):
                cfg = trackers.get_tracker(tracker_id)
                cfg["auto_fetch"] = False
                st.session_state.pending_tracker_load = cfg
                st.session_state.pending_view = BUILDER_VIEW
                st.rerun()

            if action_cols[1].button(
                "📧 Email", key=f"tracker_email_{tracker_id}",
                help="Send the consolidated report to this tracker's saved recipients",
                use_container_width=True,
            ):
                key = f"show_email_{tracker_id}"
                st.session_state[key] = not st.session_state.get(key, False)

            if action_cols[2].button(
                "📊 Generate File", key=f"tracker_dl_{tracker_id}",
                help="Runs the tracker's query fresh - this doesn't download anything by itself, "
                "it prepares the file for the 'Download file' button that appears once it's ready.",
                use_container_width=True,
            ):
                with st.spinner(
                    f"Running '{row['tracker_name']}' fresh against the database - this can take "
                    "a while for a broad date range or no client filter..."
                ):
                    cfg = trackers.get_tracker(tracker_id)
                    fresh_df, _, _ = run_tracker_query(cfg)
                    st.session_state[f"tracker_dl_bytes_{tracker_id}"] = build_report_excel(
                        fresh_df, columns=cfg.get("selected_columns")
                    )
                st.rerun()

            if row["status"] == "Active":
                if action_cols[3].button("⏸️ Inactive", key=f"tracker_deactivate_{tracker_id}", use_container_width=True):
                    trackers.set_status(tracker_id, "Inactive")
                    st.rerun()
            else:
                if action_cols[3].button("▶️ Active", key=f"tracker_activate_{tracker_id}", use_container_width=True):
                    trackers.set_status(tracker_id, "Active")
                    st.rerun()

            ready_bytes = st.session_state.get(f"tracker_dl_bytes_{tracker_id}")
            if ready_bytes:
                action_cols[4].download_button(
                    "⬇️ Download file",
                    data=ready_bytes,
                    file_name=f"{sanitize_filename(row['tracker_name'], set(), 100)}_{date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"tracker_dl_get_{tracker_id}",
                    help="Save the file that was just generated",
                    use_container_width=True,
                )
                st.success("✅ File is ready - click '⬇️ Download file' above to save it.")

            if st.session_state.get(f"show_email_{tracker_id}"):
                with st.container(border=True):
                    st.markdown("**📧 Email consolidated report**")
                    st.caption(
                        "Sends a single consolidated .xlsx (not split by anything) to the addresses below, "
                        "using this tracker's saved filters/columns run fresh right now."
                    )
                    with st.form(f"email_form_{tracker_id}", border=False):
                        col_to2, col_cc2 = st.columns(2)
                        with col_to2:
                            email_to = st.text_input(
                                "To_address", value=row["to_address"], key=f"email_to_{tracker_id}",
                                placeholder="name@example.com; separate multiple with a semicolon",
                            )
                        with col_cc2:
                            email_cc = st.text_input(
                                "CC_address", value=row["cc_address"], key=f"email_cc_{tracker_id}",
                                placeholder="optional",
                            )
                        col_save_r, col_send = st.columns(2)
                        with col_save_r:
                            save_recipients_clicked = st.form_submit_button(
                                "💾 Save recipients", use_container_width=True
                            )
                        with col_send:
                            send_now_clicked = st.form_submit_button("📧 Send Now", use_container_width=True)

                    if save_recipients_clicked:
                        trackers.update_recipients(tracker_id, email_to.strip(), email_cc.strip())
                        st.success("Recipients saved.")

                    if send_now_clicked:
                        if not email_to.strip():
                            st.error("Enter a To_address before sending.")
                        elif not outlook_available():
                            st.warning(
                                "Outlook automation (pywin32) isn't available in this environment, "
                                "so emailing is disabled here."
                            )
                        else:
                            try:
                                with st.spinner(f"Running '{row['tracker_name']}' fresh and sending..."):
                                    cfg = trackers.get_tracker(tracker_id)
                                    fresh_df, _, _ = run_tracker_query(cfg)
                                    file_bytes = build_report_excel(fresh_df, columns=cfg.get("selected_columns"))
                                    file_name = (
                                        f"{sanitize_filename(row['tracker_name'], set(), 100)}_"
                                        f"{date.today().strftime('%Y%m%d')}.xlsx"
                                    )
                                    subject = f"Case Report - {row['tracker_name']} - {date.today():%Y-%m-%d}"
                                    body = f"Please find attached the '{row['tracker_name']}' report."
                                    send_via_outlook(
                                        email_to.strip(), email_cc.strip(), subject, body, file_bytes, file_name
                                    )
                                st.success(f"✅ Sent to {email_to.strip()}.")
                            except Exception as exc:
                                st.error(f"Send failed - {exc}")

    col_prev, col_page, col_next = st.columns([1, 3, 1])
    with col_prev:
        if st.button("« Previous", disabled=st.session_state.tracker_page <= 0, use_container_width=True):
            st.session_state.tracker_page -= 1
            st.rerun()
    with col_page:
        st.caption(f"Page {st.session_state.tracker_page + 1} of {total_pages} - {len(tracker_df)} tracker(s)")
    with col_next:
        if st.button("Next »", disabled=st.session_state.tracker_page >= total_pages - 1, use_container_width=True):
            st.session_state.tracker_page += 1
            st.rerun()


client_lookup = get_client_lookup()
client_lookup["label"] = (
    client_lookup["company_name"].astype(str)
    + " (" + client_lookup["client_external_id"].astype(str) + ")"
)
label_to_id = dict(zip(client_lookup["label"], client_lookup["client_external_id"]))

# Applying a loaded tracker's saved filters has to happen before the filter
# widgets below are instantiated, since a widget only picks up a value from
# session_state[key] set *before* it renders on this run - setting it after
# (e.g. from a button handler further down the script) only takes effect on
# the *next* rerun, which is why Edit/Email Tracker actions stash the config
# here and call st.rerun() rather than trying to poke the widgets directly.
#
# Every restored value is filtered down to what's actually a valid option
# right now: a plain st.multiselect hard-errors (crashing the whole page) if
# its pre-set session_state value contains anything outside its `options`,
# which can easily happen for a tracker saved a while ago (a client renamed,
# a check retired, etc). Check Severity is the one exception - that widget
# already uses accept_new_options=True, so any saved value is valid there.
if st.session_state.pending_tracker_load is not None:
    _cfg = st.session_state.pending_tracker_load
    st.session_state.pending_tracker_load = None

    _valid_labels = [l for l in (_cfg.get("client_labels") or []) if l in label_to_id]
    st.session_state["flt_client_labels"] = _valid_labels
    _valid_client_ids = [label_to_id[l] for l in _valid_labels]

    _valid_check_names = set(
        get_check_name_lookup(tuple(sorted(_valid_client_ids)))["check_name"].dropna().tolist()
    )
    st.session_state["flt_check_names"] = [
        n for n in (_cfg.get("check_names") or []) if n in _valid_check_names
    ]

    _valid_case_statuses = set(get_case_status_lookup()["case_status"].dropna().tolist())
    st.session_state["flt_case_statuses"] = [
        s for s in (_cfg.get("case_statuses") or []) if s in _valid_case_statuses
    ]

    _valid_check_statuses = set(get_check_status_lookup()["check_status"].dropna().tolist())
    st.session_state["flt_check_statuses"] = [
        s for s in (_cfg.get("check_statuses") or []) if s in _valid_check_statuses
    ]

    st.session_state["flt_check_severities"] = _cfg.get("check_severities") or []
    st.session_state["flt_from_date"] = date.today() - timedelta(days=int(_cfg.get("lookback_days", 30)))
    st.session_state["flt_to_date"] = date.today()
    st.session_state.loaded_tracker_id = _cfg.get("tracker_id")
    st.session_state.loaded_tracker_columns = _cfg.get("selected_columns")
    st.session_state.loaded_tracker_name = _cfg.get("tracker_name", "")
    st.session_state.loaded_tracker_description = _cfg.get("tracker_description", "")
    st.session_state.loaded_tracker_to_address = _cfg.get("to_address", "")
    st.session_state.loaded_tracker_cc_address = _cfg.get("cc_address", "")
    st.session_state.loaded_tracker_schedule_frequency = _cfg.get("schedule_frequency", "")
    st.session_state.loaded_tracker_schedule_enabled = bool(st.session_state.loaded_tracker_schedule_frequency)
    st.session_state.loaded_tracker_schedule_days = _cfg.get("schedule_days") or []
    _saved_time = _cfg.get("schedule_time")
    st.session_state.loaded_tracker_schedule_time = (
        time(*(int(p) for p in _saved_time.split(":"))) if _saved_time else time(9, 0)
    )

    if _cfg.get("auto_fetch"):
        with st.spinner("Running tracker query against the database..."):
            _df, _flex_map, _distinct_client_ids = run_tracker_query(_cfg)
            _selected_columns, _available_columns = _merge_column_lists(list(_df.columns))
            if st.session_state.loaded_tracker_columns is not None:
                _saved_cols = [c for c in st.session_state.loaded_tracker_columns if c in _df.columns]
                if _saved_cols:
                    _selected_columns = _saved_cols
                    _available_columns = [c for c in _df.columns if c not in _saved_cols]
                st.session_state.loaded_tracker_columns = None
            st.session_state.selected_columns = _selected_columns
            st.session_state.available_columns = _available_columns
            st.session_state.sorter_version += 1
            _excel_bytes, _zip_bytes = _rebuild_exports(_df, _selected_columns, _flex_map)

            st.session_state.report_df = _df
            st.session_state.flex_map = _flex_map
            st.session_state.distinct_client_ids = _distinct_client_ids
            st.session_state.applied_columns = tuple(_selected_columns)
            st.session_state.report_excel_bytes = _excel_bytes
            st.session_state.report_zip_bytes = _zip_bytes
            st.session_state.antecedent_columns = [
                item["field_name"] for item in (_cfg.get("antecedent_fields") or [])
            ]
            st.session_state.antecedent_field_types = {
                item["field_name"]: item["data_type"] for item in (_cfg.get("antecedent_fields") or [])
            }
            for _key in GROUP_SPLIT_SESSION_KEYS:
                st.session_state[_key] = None

view = st.segmented_control("View", [BUILDER_VIEW, TRACKERS_VIEW], key="view_switch")

if view == TRACKERS_VIEW:
    render_saved_trackers_view()
    st.stop()

st.subheader("Filters")

col_from, col_to = st.columns(2)
with col_from:
    from_date = st.date_input("From (Case Received Date)", key="flt_from_date")
with col_to:
    to_date = st.date_input("To (Case Received Date)", key="flt_to_date")

selected_labels = st.multiselect(
    "Company Name / Client External ID (leave empty for all clients)",
    options=list(label_to_id.keys()),
    placeholder="Search by company name or client external ID...",
    key="flt_client_labels",
)
selected_client_ids = [label_to_id[label] for label in selected_labels]

check_name_lookup = get_check_name_lookup(tuple(sorted(selected_client_ids)))
selected_check_names = st.multiselect(
    "Unique Check Name (leave empty for all check names)",
    options=check_name_lookup["check_name"].dropna().tolist(),
    placeholder="Search by unique check name...",
    key="flt_check_names",
)

case_status_lookup = get_case_status_lookup()
check_status_lookup = get_check_status_lookup()

col_case_status, col_check_status, col_check_severity = st.columns(3)
with col_case_status:
    selected_case_statuses = st.multiselect(
        "Case Status (leave empty for all)",
        options=case_status_lookup["case_status"].dropna().tolist(),
        key="flt_case_statuses",
    )
with col_check_status:
    selected_check_statuses = st.multiselect(
        "Check Status (leave empty for all)",
        options=check_status_lookup["check_status"].dropna().tolist(),
        key="flt_check_statuses",
    )
with col_check_severity:
    # check_severity is a free-text legacy column with ~140 messy distinct
    # values overall and no fast way to list them unscoped (see queries.py).
    # Scoped to the selected client(s) it's quick and clean, so load it
    # automatically there, same as the check-name lookup above; with no
    # client selected, fall back to typing a value or an explicit slow load.
    if selected_client_ids:
        check_severity_lookup = get_check_severity_lookup(tuple(sorted(selected_client_ids)))
        selected_check_severities = st.multiselect(
            "Check Severity (leave empty for all)",
            options=check_severity_lookup["check_severity"].dropna().tolist(),
            accept_new_options=True,
            placeholder="Search or type a severity value...",
            key="flt_check_severities",
        )
    elif st.session_state.check_severity_options is not None:
        selected_check_severities = st.multiselect(
            "Check Severity (leave empty for all)",
            options=st.session_state.check_severity_options,
            accept_new_options=True,
            placeholder="Search or type a severity value...",
            key="flt_check_severities",
        )
    else:
        selected_check_severities = st.multiselect(
            "Check Severity (leave empty for all)",
            options=[],
            accept_new_options=True,
            placeholder="Select a company above, or type a severity value...",
            key="flt_check_severities",
        )
        if st.button("Load all severity options (~1 min, slow)"):
            with st.spinner("Scanning for distinct severity values across all clients - this can take a minute..."):
                lookup = get_check_severity_lookup()
            st.session_state.check_severity_options = lookup["check_severity"].dropna().tolist()
            st.rerun()

if from_date > to_date:
    st.error("'From' date must be on or before 'To' date.")
elif not selected_client_ids and (to_date - from_date).days > LARGE_RANGE_NO_FILTER_DAYS:
    st.warning(
        f"No client selected and the date range spans {(to_date - from_date).days} days - "
        "this query joins many large tables and may take a while. Consider narrowing the "
        "date range or picking specific clients."
    )

fetch_disabled = from_date > to_date

if st.button("📥 Fetch Report", disabled=fetch_disabled, use_container_width=True):
    with st.spinner("Running query against the database..."):
        try:
            client_filter_clause = CLIENT_FILTER_CLAUSE if selected_client_ids else ""
            check_name_filter_clause = CHECK_NAME_FILTER_CLAUSE if selected_check_names else ""
            case_status_filter_clause = CASE_STATUS_FILTER_CLAUSE if selected_case_statuses else ""
            check_status_filter_clause = CHECK_STATUS_FILTER_CLAUSE if selected_check_statuses else ""
            check_severity_filter_clause = CHECK_SEVERITY_FILTER_CLAUSE if selected_check_severities else ""
            query = CASE_REPORT_QUERY.format(
                client_filter_clause=client_filter_clause,
                check_name_filter_clause=check_name_filter_clause,
                case_status_filter_clause=case_status_filter_clause,
                check_status_filter_clause=check_status_filter_clause,
                check_severity_filter_clause=check_severity_filter_clause,
            )
            params = {"from_date": from_date, "to_date": to_date}
            if selected_client_ids:
                params["client_ids"] = selected_client_ids
            if selected_check_names:
                params["check_names"] = selected_check_names
            if selected_case_statuses:
                params["case_statuses"] = selected_case_statuses
            if selected_check_statuses:
                params["check_statuses"] = selected_check_statuses
            if selected_check_severities:
                params["check_severities"] = selected_check_severities

            df = run_query_params(query, params)

            distinct_client_ids = [
                int(cid) for cid in df["client_external_id"].dropna().unique().tolist()
            ] if not df.empty else []
            flex_map = load_flex_field_map(distinct_client_ids) if distinct_client_ids else {}

            selected_columns, available_columns = _merge_column_lists(list(df.columns))
            st.session_state.selected_columns = selected_columns
            st.session_state.available_columns = available_columns
            st.session_state.sorter_version += 1
            excel_bytes, zip_bytes = _rebuild_exports(df, selected_columns, flex_map)

            st.session_state.report_df = df
            st.session_state.flex_map = flex_map
            st.session_state.distinct_client_ids = distinct_client_ids
            st.session_state.applied_columns = tuple(selected_columns)
            st.session_state.report_excel_bytes = excel_bytes
            st.session_state.report_zip_bytes = zip_bytes
            st.session_state.antecedent_columns = []
            st.session_state.antecedent_field_types = {}
            for _key in GROUP_SPLIT_SESSION_KEYS:
                st.session_state[_key] = None
            st.success(f"Fetched {len(df)} row(s) across {len(distinct_client_ids)} client(s).")
        except Exception as exc:
            st.session_state.report_df = None
            st.session_state.report_excel_bytes = None
            st.session_state.report_zip_bytes = None
            st.error(f"Query failed: {exc}")

st.divider()

if st.session_state.report_df is not None:
    df = st.session_state.report_df

    st.subheader("Add Antecedent Fields")
    st.session_state.enable_antecedents = st.checkbox(
        "Enable antecedent fields for this report",
        value=st.session_state.enable_antecedents,
        help=(
            "Off by default so fetching a report never pays for this. Pulls specific "
            "per-check custom fields (e.g. Employment/Education antecedents like "
            "Designation or Salary) in as extra columns, mapped by Case_Check_id. "
            "Kept separate from the main query since joining it directly would "
            "multiply every check's row once per field it has data for."
        ),
    )

    if st.session_state.antecedent_columns:
        st.markdown("**Added antecedent field(s)**")
        for _added_name in list(st.session_state.antecedent_columns):
            col_added_name, col_added_remove = st.columns([5, 1])
            with col_added_name:
                _added_type = st.session_state.antecedent_field_types.get(_added_name, "verified_data")
                st.caption(f"{_added_name} · {'Verified Data' if _added_type == 'verified_data' else 'Stated Data'}")
            with col_added_remove:
                if st.button("🗑️ Remove", key=f"remove_antecedent_{_added_name}", use_container_width=True):
                    df = df.drop(columns=[_added_name]) if _added_name in df.columns else df
                    st.session_state.report_df = df
                    st.session_state.antecedent_columns.remove(_added_name)
                    st.session_state.antecedent_field_types.pop(_added_name, None)
                    if _added_name in st.session_state.selected_columns:
                        st.session_state.selected_columns.remove(_added_name)
                    if _added_name in st.session_state.available_columns:
                        st.session_state.available_columns.remove(_added_name)
                    st.session_state.sorter_version += 1
                    st.rerun()

    if st.session_state.enable_antecedents:
        case_check_ids = df["Case_Check_id"].dropna().unique().tolist()
        with st.spinner("Loading available antecedent field names..."):
            field_name_lookup = get_field_name_lookup(tuple(sorted(case_check_ids)))

        col_field_names, col_data_type = st.columns([3, 1])
        with col_field_names:
            selected_field_names = st.multiselect(
                "Antecedent Field(s)",
                options=field_name_lookup["field_name"].dropna().tolist(),
                placeholder="Search for a field name (e.g. Designation, Salary)...",
            )
        with col_data_type:
            data_type_label = st.radio("Value to pull", options=["Verified Data", "Stated Data"])
        data_column = "verified_data" if data_type_label == "Verified Data" else "stated_data"

        duplicate_field_names = [n for n in selected_field_names if n in st.session_state.antecedent_columns]
        if duplicate_field_names:
            st.warning(
                f"Already added: {', '.join(duplicate_field_names)}. Use 'Remove' above first if you meant to "
                "replace one instead of adding it again."
            )

        if st.button(
            "➕ Add antecedent column(s)",
            disabled=not selected_field_names,
            use_container_width=True,
        ):
            new_field_names = [n for n in selected_field_names if n not in st.session_state.antecedent_columns]
            if not new_field_names:
                st.warning("Nothing to add - every selected field is already added.")
            else:
                with st.spinner("Fetching antecedent field data..."):
                    antecedent_df = fetch_antecedent_columns(case_check_ids, new_field_names, data_column)

                    df = df.drop(columns=[c for c in new_field_names if c in df.columns])
                    df = df.merge(antecedent_df, how="left", left_on="Case_Check_id", right_index=True)
                    for name in new_field_names:
                        if name not in df.columns:
                            df[name] = pd.NA

                    st.session_state.report_df = df
                    for name in new_field_names:
                        if name in st.session_state.available_columns:
                            st.session_state.available_columns.remove(name)
                        if name not in st.session_state.selected_columns:
                            st.session_state.selected_columns.append(name)
                        if name not in st.session_state.antecedent_columns:
                            st.session_state.antecedent_columns.append(name)
                        st.session_state.antecedent_field_types[name] = data_column
                    st.session_state.sorter_version += 1
                st.rerun()

    st.divider()

    st.subheader("Customize Columns")

    with st.expander("ℹ️ Which columns are Case-wise vs Check-wise?", expanded=False):
        st.caption(
            "This report is one row per (case, check) pair. Case-wise columns "
            "are anchored to the case's ARS number (case_ars_no) and repeat "
            "across every check row of the same case; Check-wise columns are "
            "anchored to Case_Check_id and vary row-by-row even within one "
            "case. Hiding every Check-wise column below (starting with "
            "Case_Check_id itself) collapses the report from check-level to "
            "case-level, since the now-duplicate rows merge into one."
        )
        present_case_cols = [c for c in df.columns if c in CASE_LEVEL_COLUMNS]
        present_check_cols = [c for c in df.columns if c in CHECK_LEVEL_COLUMNS]
        present_other_cols = [
            c for c in df.columns if c not in CASE_LEVEL_COLUMNS and c not in CHECK_LEVEL_COLUMNS
        ]
        col_case, col_check = st.columns(2)
        with col_case:
            st.markdown(f"**Case-wise · by ARS number ({len(present_case_cols)})**")
            st.caption(", ".join(present_case_cols) or "-")
        with col_check:
            st.markdown(f"**Check-wise · by Case_Check_id ({len(present_check_cols)})**")
            st.caption(", ".join(present_check_cols) or "-")
        if present_other_cols:
            st.markdown(f"**Antecedent fields ({len(present_other_cols)})**")
            st.caption(
                "Per-check custom fields added via 'Add Antecedent Fields' above - "
                "these vary per check, so behave like Check-wise columns."
            )
            st.caption(", ".join(present_other_cols))

    st.caption(
        "Drag columns between the two lists to show/hide them, and drag within "
        "'Selected' to reorder. This works on the raw column names - the "
        "split-by-client download still renames its own flex fields afterwards. "
        "Hiding a column that distinguishes otherwise-identical rows (e.g. "
        "Case_Check_id, to move from check-level to case-level) automatically "
        "collapses those rows into one, just like hiding a field in an Excel "
        "pivot table."
    )

    btn_all, btn_none, btn_reset = st.columns(3)
    with btn_all:
        if st.button("Select all", use_container_width=True):
            st.session_state.selected_columns += st.session_state.available_columns
            st.session_state.available_columns = []
            st.session_state.sorter_version += 1
            st.rerun()
    with btn_none:
        if st.button("Deselect all", use_container_width=True):
            st.session_state.available_columns += st.session_state.selected_columns
            st.session_state.selected_columns = []
            st.session_state.sorter_version += 1
            st.rerun()
    with btn_reset:
        if st.button("Reset order", use_container_width=True):
            original_order = list(df.columns)
            st.session_state.selected_columns = [
                c for c in original_order if c in st.session_state.selected_columns
            ]
            st.session_state.available_columns = [
                c for c in original_order if c in st.session_state.available_columns
            ]
            st.session_state.sorter_version += 1
            st.rerun()

    sorted_containers = sort_items(
        [
            {"header": "Available (hidden)", "items": st.session_state.available_columns},
            {"header": "Selected (shown, in this order)", "items": st.session_state.selected_columns},
        ],
        multi_containers=True,
        direction="vertical",
        key=f"col_sorter_{st.session_state.sorter_version}",
    )
    st.session_state.available_columns = sorted_containers[0]["items"]
    st.session_state.selected_columns = sorted_containers[1]["items"]

    selected_columns = st.session_state.selected_columns
    total_columns = len(selected_columns) + len(st.session_state.available_columns)
    st.caption(f"{len(selected_columns)} of {total_columns} columns selected.")

    if not selected_columns:
        st.warning("Select at least one column before downloading.")
    elif tuple(selected_columns) != st.session_state.applied_columns:
        st.info("Column selection changed - click 'Apply column selection' below to refresh the downloads.")

    if st.button(
        "🔄 Apply column selection to downloads",
        disabled=not selected_columns,
        use_container_width=True,
    ):
        with st.spinner("Rebuilding exports..."):
            excel_bytes, zip_bytes = _rebuild_exports(df, selected_columns, st.session_state.flex_map)
            st.session_state.report_excel_bytes = excel_bytes
            st.session_state.report_zip_bytes = zip_bytes
            st.session_state.applied_columns = tuple(selected_columns)
        st.rerun()

    st.divider()

    st.subheader("Report Preview")
    preview_columns = selected_columns or list(df.columns)
    preview_df = aggregate_selected_columns(df[preview_columns])
    st.caption(
        f"{len(preview_df)} rows x {len(preview_df.columns)} columns shown "
        "(reflects your current column selection/order, aggregated to that level "
        "of detail; raw column names shown below - the split-by-client download "
        "aggregates within each client's sheet separately)"
    )
    st.dataframe(preview_df, use_container_width=True)

    col_consolidated, col_split = st.columns(2)
    with col_consolidated:
        st.download_button(
            label="📊 Download Consolidated Report (.xlsx)",
            data=st.session_state.report_excel_bytes,
            file_name=f"Client_Case_Report_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(f"Single sheet, {len(st.session_state.applied_columns)} column(s) per your applied selection.")
    with col_split:
        st.download_button(
            label="🗂️ Download Split Report (.zip)",
            data=st.session_state.report_zip_bytes,
            file_name=f"Client_Case_Report_Split_{date.today().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            use_container_width=True,
        )
        st.caption("One .xlsx per company, columns per your applied selection with each client's own flex field names.")

    st.divider()
    with st.expander("💾 Save this filter set as a Tracker", expanded=True):
        if st.session_state.loaded_tracker_id:
            col_editing_caption, col_stop_editing = st.columns([4, 1])
            with col_editing_caption:
                st.caption(
                    f"Editing tracker #{st.session_state.loaded_tracker_id} - 'Update Tracker' saves changes "
                    "to it, or use 'Save as New Tracker' to keep the original and create a copy instead."
                )
            with col_stop_editing:
                if st.button("✖ Stop editing", use_container_width=True):
                    st.session_state.loaded_tracker_id = None
                    st.session_state.loaded_tracker_name = ""
                    st.session_state.loaded_tracker_description = ""
                    st.session_state.loaded_tracker_to_address = ""
                    st.session_state.loaded_tracker_cc_address = ""
                    st.session_state.loaded_tracker_schedule_enabled = False
                    st.session_state.loaded_tracker_schedule_frequency = ""
                    st.session_state.loaded_tracker_schedule_days = []
                    st.session_state.loaded_tracker_schedule_time = time(9, 0)
                    st.rerun()

        # Frequency lives outside the form below (and Day(s)/Time alongside it,
        # to keep the row together) specifically so picking a Frequency reacts
        # immediately: a plain st.form only pushes any of its contained widgets'
        # values on submit, so Day(s)'s visibility couldn't depend on Frequency's
        # live value if Frequency were inside the same form - the whole form
        # would need submitting before Day(s) could even appear.
        st.markdown("**Schedule** (optional)")
        schedule_enabled_input = st.checkbox(
            "Schedule this tracker to send automatically",
            value=st.session_state.loaded_tracker_schedule_enabled,
            key="tracker_schedule_enabled_input",
        )
        _real_frequencies = [f for f in trackers.SCHEDULE_FREQUENCIES if f]
        if schedule_enabled_input:
            tracker_freq_input = st.selectbox(
                "Frequency",
                options=_real_frequencies,
                index=_real_frequencies.index(st.session_state.loaded_tracker_schedule_frequency)
                if st.session_state.loaded_tracker_schedule_frequency in _real_frequencies
                else 0,
                key="tracker_freq_input",
            )
            if tracker_freq_input == "Weekly":
                col_days, col_time = st.columns([2, 1])
                with col_days:
                    tracker_days_input = st.multiselect(
                        "Day(s)",
                        options=trackers.WEEKDAY_NAMES,
                        default=st.session_state.loaded_tracker_schedule_days,
                        key="tracker_days_input",
                    )
                with col_time:
                    tracker_time_input = st.time_input(
                        "Time", value=st.session_state.loaded_tracker_schedule_time, key="tracker_time_input"
                    )
            else:
                tracker_days_input = []
                tracker_time_input = st.time_input(
                    "Time", value=st.session_state.loaded_tracker_schedule_time, key="tracker_time_input"
                )
            st.caption(
                "A separate scheduler script (scheduled_tracker_runner.py) has to be running/scheduled on this "
                "machine (e.g. via Windows Task Scheduler) for this to actually fire - saving a schedule here "
                "only stores the setting, it doesn't start anything by itself."
            )
        else:
            tracker_freq_input = ""
            tracker_days_input = []
            tracker_time_input = st.session_state.loaded_tracker_schedule_time

        # Name/Description/Recipients + the save buttons live inside one form so
        # a submit atomically captures whatever's currently entered. Without a
        # form, a text_input only pushes its value to session_state on blur, sent
        # as its own message right before - but not necessarily strictly before -
        # the button-click message; clicking Save right after typing (fast
        # typist, or anything scripted) could race and save a stale/blank value.
        # A form has no such race: nothing is sent until submit, and the submit
        # bundles every contained widget's current value in one message.
        with st.form("save_tracker_form", border=False):
            col_tname, col_tdesc = st.columns(2)
            with col_tname:
                tracker_name_input = st.text_input(
                    "Tracker Name", value=st.session_state.loaded_tracker_name, key="tracker_name_input"
                )
            with col_tdesc:
                tracker_desc_input = st.text_input(
                    "Tracker Description", value=st.session_state.loaded_tracker_description, key="tracker_desc_input"
                )
            st.caption(
                f"Date range is saved as a rolling {(to_date - from_date).days}-day window (today minus "
                f"{(to_date - from_date).days} days), not these exact dates - so re-running it later always "
                "pulls fresh, recent data instead of replaying a range that's aged out."
            )

            st.markdown("**Recipients** (used by the consolidated-report email in Saved Trackers)")
            col_to, col_cc = st.columns(2)
            with col_to:
                tracker_to_input = st.text_input(
                    "To_address", value=st.session_state.loaded_tracker_to_address, key="tracker_to_input",
                    placeholder="name@example.com; separate multiple with a semicolon",
                )
            with col_cc:
                tracker_cc_input = st.text_input(
                    "CC_address", value=st.session_state.loaded_tracker_cc_address, key="tracker_cc_input",
                    placeholder="optional",
                )

            col_save, col_save_new = st.columns(2)
            with col_save:
                save_label = "💾 Update Tracker" if st.session_state.loaded_tracker_id else "💾 Save Tracker"
                save_clicked = st.form_submit_button(save_label, use_container_width=True)
            with col_save_new:
                save_new_clicked = (
                    st.form_submit_button("💾 Save as New Tracker", use_container_width=True)
                    if st.session_state.loaded_tracker_id
                    else False
                )

        if save_clicked or save_new_clicked:
            tracker_name = tracker_name_input.strip()
            if not tracker_name:
                st.error("Enter a Tracker Name before saving.")
            elif tracker_freq_input == "Weekly" and not tracker_days_input:
                st.error("Pick at least one day for a Weekly schedule, or uncheck 'Schedule this tracker'.")
            else:
                tracker_cfg = {
                    "tracker_name": tracker_name,
                    "tracker_description": tracker_desc_input.strip(),
                    "client_ids": selected_client_ids,
                    "client_labels": selected_labels,
                    "lookback_days": (to_date - from_date).days,
                    "check_names": selected_check_names,
                    "case_statuses": selected_case_statuses,
                    "check_statuses": selected_check_statuses,
                    "check_severities": selected_check_severities,
                    "selected_columns": st.session_state.selected_columns,
                    "antecedent_fields": [
                        {
                            "field_name": name,
                            "data_type": st.session_state.antecedent_field_types.get(name, "verified_data"),
                        }
                        for name in st.session_state.antecedent_columns
                    ],
                    "to_address": tracker_to_input.strip(),
                    "cc_address": tracker_cc_input.strip(),
                    "schedule_frequency": tracker_freq_input,
                    "schedule_days": tracker_days_input if tracker_freq_input == "Weekly" else [],
                    "schedule_time": tracker_time_input.strftime("%H:%M"),
                }
                existing_id = None if save_new_clicked else st.session_state.loaded_tracker_id
                st.session_state.loaded_tracker_id = trackers.save_tracker(tracker_cfg, tracker_id=existing_id)
                st.session_state.loaded_tracker_name = tracker_name
                st.session_state.loaded_tracker_description = tracker_cfg["tracker_description"]
                st.session_state.loaded_tracker_to_address = tracker_cfg["to_address"]
                st.session_state.loaded_tracker_cc_address = tracker_cfg["cc_address"]
                st.session_state.loaded_tracker_schedule_enabled = schedule_enabled_input
                st.session_state.loaded_tracker_schedule_frequency = tracker_cfg["schedule_frequency"]
                st.session_state.loaded_tracker_schedule_days = tracker_cfg["schedule_days"]
                st.session_state.loaded_tracker_schedule_time = tracker_time_input
                verb = "Saved as a new tracker" if save_new_clicked else "Saved"
                st.success(f"{verb}: '{tracker_name}'. Find it on the '{TRACKERS_VIEW}' view.")

    st.divider()

    st.subheader("More Ways to Split")
    st.caption(
        "Each of these regroups your current column selection by a different "
        "column instead of by client - built on demand from the report already "
        "in memory (no extra database query), and refused up front rather than "
        f"built if the column has more than {MAX_SPLIT_GROUPS} distinct values, "
        "since that's almost always the wrong column to split by."
    )

    st.markdown("**Split by Process Name**")
    _render_group_split_download(
        "Process Name", "Process Name", "Process_name", "split_by_process_zip", df, selected_columns
    )

    st.markdown("**Split by Location**")
    _render_group_split_download("Location", "Location", "location", "split_by_location_zip", df, selected_columns)

    st.markdown("**Split by Case Flex Field**")
    flex_field_options = [
        c for c in (f"CASE_FLEX_FIELD{i}" for i in range(1, 31))
        if c in df.columns and df[c].astype(str).str.strip().replace("nan", "").ne("").any()
    ]
    if flex_field_options:
        selected_flex_field = st.selectbox(
            "Flex field to split by",
            options=flex_field_options,
            help="Only flex fields with at least one non-empty value in this report are listed.",
            label_visibility="collapsed",
        )
        _render_group_split_download(
            f"Flex Field ({selected_flex_field})",
            "Case Flex Field",
            selected_flex_field,
            "split_by_flexfield_zip",
            df,
            selected_columns,
        )
    else:
        st.caption("No Case Flex Fields have data in this report.")

    st.markdown("**Split by Antecedent Field**")
    if st.session_state.antecedent_columns:
        antecedent_split_options = [c for c in st.session_state.antecedent_columns if c in df.columns]
        selected_antecedent_field = st.selectbox(
            "Antecedent field to split by",
            options=antecedent_split_options,
            label_visibility="collapsed",
        )
        _render_group_split_download(
            f"Antecedent Field ({selected_antecedent_field})",
            "Antecedent Field",
            selected_antecedent_field,
            "split_by_antecedent_zip",
            df,
            selected_columns,
        )
    else:
        st.caption("Add an antecedent field above first (see 'Add Antecedent Fields').")
