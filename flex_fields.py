"""Per-client CASE_FLEX_FIELD1..30 -> real field name mapping.

Each client can name/reuse their 30 flexi fields differently (and can rename a
field over time, leaving old, now-"Passive" rows behind), so the mapping is
built per client_external_id from `ec_client_case_fields`.
"""

import pandas as pd

from db import run_query_params
from queries import FLEX_FIELD_QUERY

MAX_FLEX_FIELDS = 30


def load_flex_field_map(client_external_ids: list[int] | None = None) -> dict[int, dict[int, str]]:
    """Return {client_external_id: {field_index (1-30): field_name}}.

    When a client has more than one row for the same field_index (renames,
    retired fields), the row with the highest client_field_id (the most
    recently created/edited definition) wins.
    """
    query = FLEX_FIELD_QUERY
    params = {}
    if client_external_ids:
        query += "\nAND eclient.client_external_id IN :client_ids"
        params["client_ids"] = list(client_external_ids)

    df = run_query_params(query, params)
    if df.empty:
        return {}

    df = df.sort_values("client_field_id").drop_duplicates(
        subset=["client_external_id", "field_id"], keep="last"
    )

    mapping: dict[int, dict[int, str]] = {}
    for row in df.itertuples(index=False):
        client_id = row.client_external_id
        if pd.isna(client_id):
            continue
        name = (row.field_name or "").strip()
        if not name:
            continue
        mapping.setdefault(int(client_id), {})[int(row.field_id)] = name
    return mapping


def flex_column_rename_map(client_external_id, flex_map: dict[int, dict[int, str]]) -> dict[str, str]:
    """Build the {'CASE_FLEX_FIELDn': 'Real Name'} rename map for one client.

    Falls back to a readable generic name ("Flex Field n") for indices the
    client has no active mapping for, so no column is ever left with the raw
    ecff.CASE_FLEX_FIELDn name in the exported report.
    """
    client_fields = flex_map.get(int(client_external_id), {}) if pd.notna(client_external_id) else {}
    rename = {}
    for i in range(1, MAX_FLEX_FIELDS + 1):
        col = f"CASE_FLEX_FIELD{i}"
        rename[col] = client_fields.get(i, f"Flex Field {i}")
    return rename
