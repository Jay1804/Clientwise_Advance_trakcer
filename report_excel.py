"""Builds the two downloadable report formats:

- build_report_excel: a single consolidated .xlsx with every row in one sheet.
- build_split_report_zip: a .zip with one .xlsx per client, each with that
  client's own CASE_FLEX_FIELD1..30 names (see flex_fields.py) - this only
  makes sense per-file since a single flat sheet can't carry a different
  header per client.
"""

import io
import re
import zipfile

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from flex_fields import flex_column_rename_map

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
MAX_COL_WIDTH = 45
MIN_COL_WIDTH = 10

_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# A split with more groups than this is almost always someone picking a
# near-unique column (e.g. a free-text antecedent field) by mistake rather
# than a genuine reporting dimension - refuse rather than silently building
# thousands of one-row .xlsx files.
MAX_SPLIT_GROUPS = 300


class TooManySplitGroupsError(ValueError):
    """Raised by build_group_split_zip when the chosen column has too many
    distinct values to split by (see MAX_SPLIT_GROUPS)."""

    def __init__(self, group_count: int):
        self.group_count = group_count
        super().__init__(
            f"That column has {group_count} distinct values - splitting into that "
            f"many files isn't supported (limit is {MAX_SPLIT_GROUPS}). Pick a "
            "coarser column to split by."
        )


def aggregate_selected_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse rows that are identical across every given column down to one -
    the same effect as hiding a field in an Excel pivot table: e.g. hiding
    Case_Check_id collapses the several per-check rows of a case into a
    single case-level row. A no-op when the given columns are already
    granular enough that no rows are actually identical.
    """
    if df.empty or df.shape[1] == 0:
        return df

    return df.drop_duplicates()


def sanitize_filename(name: str, used: set[str], max_len: int, fallback: str = "Unknown Client") -> str:
    """Turn an arbitrary group value into a safe, deduped filename stem (no
    extension). Public since it's also used outside this module for the
    per-group email attachment names in app.py."""
    name = str(name) if name and str(name) != "nan" else fallback
    name = _INVALID_FILENAME_CHARS.sub(" ", name).strip() or fallback
    name = name[:max_len]
    base, i = name, 1
    while name.lower() in used:
        suffix = f" ({i})"
        name = base[: max_len - len(suffix)] + suffix
        i += 1
    used.add(name.lower())
    return name


def _write_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame) -> None:
    ws: Worksheet = wb.create_sheet(title=sheet_name)
    ws.append(list(df.columns))

    for row in df.itertuples(index=False):
        ws.append(["" if pd.isna(v) else v for v in row])

    n_cols = len(df.columns)
    for col in range(1, n_cols + 1):
        header_cell = ws.cell(row=1, column=col)
        header_cell.fill = HEADER_FILL
        header_cell.font = HEADER_FONT
        header_cell.alignment = HEADER_ALIGNMENT

        header_len = len(str(header_cell.value or ""))
        try:
            max_data_len = df.iloc[:, col - 1].astype(str).map(len).max()
        except (ValueError, IndexError):
            max_data_len = 0
        width = max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, max(header_len, max_data_len) + 2))
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and n_cols >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{ws.max_row}"
    ws.row_dimensions[1].height = 30


def _single_sheet_workbook_bytes(sheet_name: str, df: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    _write_sheet(wb, sheet_name, df)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _prepare_client_sheet(
    client_df: pd.DataFrame, client_external_id, flex_map: dict, columns: list[str] | None = None
) -> pd.DataFrame:
    if columns is not None:
        client_df = client_df[[c for c in columns if c in client_df.columns]]

    rename = flex_column_rename_map(client_external_id, flex_map)
    client_df = client_df.rename(columns=rename)

    # Drop flex columns that fell back to a generic "Flex Field n" name AND
    # are empty for every row in this client's data - nothing to show.
    fallback_cols = [v for v in rename.values() if v.startswith("Flex Field ")]
    for col in fallback_cols:
        if col in client_df.columns and client_df[col].map(lambda v: pd.isna(v) or str(v).strip() == "").all():
            client_df = client_df.drop(columns=[col])

    return aggregate_selected_columns(client_df)


def build_report_excel(df: pd.DataFrame, columns: list[str] | None = None) -> bytes:
    """Single consolidated workbook, one sheet, raw column names.

    `columns`, if given, selects and orders which columns are written -
    used to honor the user's column show/hide + reorder choices in the UI.
    Hiding a column that distinguishes otherwise-identical rows (e.g.
    Case_Check_id) collapses those rows into one - see `aggregate_selected_columns`.
    """
    if columns is not None:
        df = df[[c for c in columns if c in df.columns]]

    if df.empty:
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet(title="Report")
        ws.append(["No rows matched the selected filters."])
        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    return _single_sheet_workbook_bytes("Report", aggregate_selected_columns(df))


def build_split_report_zip(df: pd.DataFrame, flex_map: dict, columns: list[str] | None = None) -> bytes:
    """Zip of one .xlsx per client (by Company_name), each with that
    client's own CASE_FLEX_FIELD1-30 names applied via flex_map.

    `columns` (raw column names, pre-rename) selects/orders the output the
    same way as `build_report_excel`; grouping always uses the full `df` so
    it still works even if the grouping columns themselves are hidden.
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        if df.empty:
            zf.writestr("No rows matched the selected filters.txt", "No rows matched the selected filters.")
        else:
            used_names: set[str] = set()
            group_cols = ["client_external_id", "Company_name"]
            for (client_external_id, company_name), client_df in df.groupby(group_cols, dropna=False, sort=True):
                sheet_df = _prepare_client_sheet(client_df, client_external_id, flex_map, columns=columns)
                file_bytes = _single_sheet_workbook_bytes("Report", sheet_df)
                file_name = sanitize_filename(company_name, used_names, max_len=150) + ".xlsx"
                zf.writestr(file_name, file_bytes)

    return buffer.getvalue()


def build_group_split_zip(
    df: pd.DataFrame, group_column: str, columns: list[str] | None = None
) -> bytes:
    """Zip of one .xlsx per distinct value of `group_column` (e.g. Process_name,
    location, a Case Flex Field, or an added antecedent field) - the same
    column selection/aggregation as build_report_excel, just grouped by an
    arbitrary column instead of by client.

    Memory stays bounded regardless of report size: pandas' groupby produces
    one group's rows at a time as the loop advances (it doesn't materialize
    every group up front), and each group's workbook bytes are written
    straight into the open zip and discarded before the next group starts -
    at most one group's DataFrame and one workbook are ever live at once.
    Raises TooManySplitGroupsError up front (before building anything) if the
    column has more distinct values than MAX_SPLIT_GROUPS.
    """
    if not df.empty and group_column in df.columns:
        group_count = df[group_column].nunique(dropna=False)
        if group_count > MAX_SPLIT_GROUPS:
            raise TooManySplitGroupsError(group_count)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        if df.empty or group_column not in df.columns:
            zf.writestr("No rows matched the selected filters.txt", "No rows matched the selected filters.")
        else:
            used_names: set[str] = set()
            for group_value, group_df in df.groupby(group_column, dropna=False, sort=True):
                sheet_df = group_df[[c for c in columns if c in group_df.columns]] if columns is not None else group_df
                sheet_df = aggregate_selected_columns(sheet_df)
                file_bytes = _single_sheet_workbook_bytes("Report", sheet_df)
                file_name = sanitize_filename(group_value, used_names, max_len=150, fallback="(Blank)") + ".xlsx"
                zf.writestr(file_name, file_bytes)

    return buffer.getvalue()


def build_single_group_excel(
    df: pd.DataFrame, group_column: str, group_value, columns: list[str] | None = None
) -> bytes:
    """The single-group counterpart to build_group_split_zip: just the one
    workbook for `group_value`, used when emailing one split file at a time
    instead of generating/unzipping the whole batch."""
    if pd.isna(group_value):
        mask = df[group_column].isna()
    else:
        mask = df[group_column] == group_value
    group_df = df[mask]
    if columns is not None:
        group_df = group_df[[c for c in columns if c in group_df.columns]]
    return _single_sheet_workbook_bytes("Report", aggregate_selected_columns(group_df))


def build_mapping_template_excel(dimension_label: str, values: list) -> bytes:
    """Downloadable starter file for a split-by-X email mapping upload: one
    row per distinct value already present in the report for that dimension,
    with blank To_address/CC_address columns for the user to fill in."""
    template_df = pd.DataFrame({
        dimension_label: values,
        "To_address": [""] * len(values),
        "CC_address": [""] * len(values),
    })
    return _single_sheet_workbook_bytes("Mapping", template_df)
