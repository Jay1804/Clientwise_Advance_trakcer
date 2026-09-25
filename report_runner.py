"""Runs CASE_REPORT_QUERY (+ antecedent fields) from a plain config dict.

Deliberately Streamlit-free (no `st.*` calls, no import of `app.py`) so it can
be imported both by the interactive app and by `scheduled_tracker_runner.py`,
a plain script with no Streamlit runtime around it.
"""

from datetime import date, timedelta

import pandas as pd

from db import run_query_params
from flex_fields import load_flex_field_map
from queries import (
    ANTECEDENT_DATA_QUERY,
    CASE_REPORT_QUERY,
    CASE_STATUS_FILTER_CLAUSE,
    CHECK_NAME_FILTER_CLAUSE,
    CHECK_SEVERITY_FILTER_CLAUSE,
    CHECK_STATUS_FILTER_CLAUSE,
    CLIENT_FILTER_CLAUSE,
)


def fetch_antecedent_columns(case_check_ids: list, field_names: list[str], data_column: str) -> pd.DataFrame:
    """One column per field_name, indexed by case_check_id, holding that
    field's stated_data or verified_data. A case_check_id/field_name pair
    should be unique but isn't DB-enforced, so ties are broken by keeping the
    most recently updated row, same as flex_fields.py does for flex columns."""
    if not case_check_ids or not field_names:
        return pd.DataFrame()

    raw = run_query_params(
        ANTECEDENT_DATA_QUERY,
        {"field_names": field_names, "case_check_ids": case_check_ids},
    )
    if raw.empty:
        return pd.DataFrame()

    raw = raw.sort_values("data_updated_on").drop_duplicates(
        subset=["case_check_id", "field_name"], keep="last"
    )
    pivoted = raw.pivot(index="case_check_id", columns="field_name", values=data_column)
    return pivoted.reindex(columns=field_names)


def merge_antecedent_fields_into(df: pd.DataFrame, antecedent_fields: list[dict]) -> pd.DataFrame:
    """Add/replace one column per {field_name, data_type} entry (grouped by
    data_type so each distinct value type is fetched in one query), mapped by
    Case_Check_id - the same merge the interactive 'Add antecedent column(s)'
    button does, reused here so a saved tracker's antecedent choices are
    restored automatically when it's re-run."""
    if not antecedent_fields or df.empty:
        return df

    case_check_ids = df["Case_Check_id"].dropna().unique().tolist()
    by_type: dict[str, list[str]] = {}
    for item in antecedent_fields:
        by_type.setdefault(item["data_type"], []).append(item["field_name"])

    for data_column, field_names in by_type.items():
        antecedent_df = fetch_antecedent_columns(case_check_ids, field_names, data_column)
        df = df.drop(columns=[c for c in field_names if c in df.columns])
        df = df.merge(antecedent_df, how="left", left_on="Case_Check_id", right_index=True)
        for name in field_names:
            if name not in df.columns:
                df[name] = pd.NA
    return df


def run_tracker_query(cfg: dict) -> tuple[pd.DataFrame, dict, list[int]]:
    """Runs CASE_REPORT_QUERY (plus any saved antecedent fields) from a plain
    config dict - the same query the interactive Fetch Report button builds
    from live widget values, but reusable for a saved tracker's Download/Email/
    scheduled-send actions. `cfg['lookback_days']` is resolved against *today*,
    not a frozen date, so a tracker re-run next month still pulls a fresh
    recent window."""
    client_ids = cfg.get("client_ids") or []
    check_names = cfg.get("check_names") or []
    case_statuses = cfg.get("case_statuses") or []
    check_statuses = cfg.get("check_statuses") or []
    check_severities = cfg.get("check_severities") or []
    to_date = date.today()
    from_date = to_date - timedelta(days=int(cfg.get("lookback_days", 30)))

    query = CASE_REPORT_QUERY.format(
        client_filter_clause=CLIENT_FILTER_CLAUSE if client_ids else "",
        check_name_filter_clause=CHECK_NAME_FILTER_CLAUSE if check_names else "",
        case_status_filter_clause=CASE_STATUS_FILTER_CLAUSE if case_statuses else "",
        check_status_filter_clause=CHECK_STATUS_FILTER_CLAUSE if check_statuses else "",
        check_severity_filter_clause=CHECK_SEVERITY_FILTER_CLAUSE if check_severities else "",
    )
    params = {"from_date": from_date, "to_date": to_date}
    if client_ids:
        params["client_ids"] = client_ids
    if check_names:
        params["check_names"] = check_names
    if case_statuses:
        params["case_statuses"] = case_statuses
    if check_statuses:
        params["check_statuses"] = check_statuses
    if check_severities:
        params["check_severities"] = check_severities

    df = run_query_params(query, params)
    df = merge_antecedent_fields_into(df, cfg.get("antecedent_fields") or [])

    distinct_client_ids = [
        int(cid) for cid in df["client_external_id"].dropna().unique().tolist()
    ] if not df.empty else []
    flex_map = load_flex_field_map(distinct_client_ids) if distinct_client_ids else {}
    return df, flex_map, distinct_client_ids
