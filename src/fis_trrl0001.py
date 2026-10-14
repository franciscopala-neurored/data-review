"""
owner: Francisco Palá
Takes a set of .xls files (FIS sample), concatenates them into a single .csv, and compares it against a TMS .csv export.
"""

import json
import re
import shutil

import pandas as pd
from pathlib import Path
from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TableCellProperties, TextProperties
from odf.table import Table, TableRow, TableCell
from odf.text import P
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from compare_datasets import compare_datasets

FONT_NAME = "Aptos Narrow"
HEADER_BG = "1F4E78"
HEADER_FONT_SIZE = 14
MATCH_Y_BG = "C6EFCE"
MATCH_N_BG = "FFC7CE"


DATA_DIR = Path("./data/TRRL0001")
RAW_DIR = DATA_DIR / "raw"
DEV_DIR = DATA_DIR / "dev"
PROD_DIR = DATA_DIR / "prod"

SNAPSHOT_DIR = RAW_DIR / "09_17_2026"
FIS_RAW_XLS_DIR = SNAPSHOT_DIR / "FIS"
FIS_RAW_CSV = SNAPSHOT_DIR / "FIS.csv"
TMS_RAW_FILE = SNAPSHOT_DIR / "TMS.csv"

# Column order of the FIS "Transportation Request List" report, i.e. the raw
# FIS export schema. Used to flatten the report's 2-row header (group row +
# field row) into the same flat column names as the legacy TRRL0001_FIS.ods
# export, so the rest of the pipeline (FIS_TMS_MAP, MATCHCODE_PAIRS, ...)
# doesn't need to change. The report's first column ("Status") is a
# UI-only row-status marker with no data-model equivalent and is dropped.
FIS_COLUMNS = [
    'Dispatch Status ', 'Tmp. Sep.', 'Transportation Execution Status',
    'Shipment Maintype', 'Shipment', 'Sequence', 'MoT', 'CntrType', 'Weight',
    'Unit', 'Remark', 'ProductIndicator', 'Temp[C]', 'DG', 'DGAccept',
    'Overdimension', 'US Flag', 'P/R/I', 'for', 'ExcludeAuto Dispatch',
    'AutoDispatchStatus', 'Multistop', 'ImportExport', 'Start Locode',
    'Start Place Name', 'Start Place Supp', 'Start Date', 'Start Time',
    'PosDel Locode', 'PosDel Place Name', 'PosDel Place Supp',
    'PosDel Postcode', 'PosDel City', 'PosDel Date', 'PosDel Time Sign',
    'PosDel Time', 'End Locode', 'End Place Name', 'End Place Supp',
    'End Date', 'End Time', 'Container Number', 'Container Rel.',
    'HandoverReference', 'LastFree Date', 'Number', 'SubconName',
    'SubconSupp', 'Status', 'GroupingNo.', 'Number.1', 'SubconName.1',
    'SubconSupp.1', 'Approval', 'Announced Delay',
]

FIS_DEV_FILE = DEV_DIR / "TRRL0001_FIS.csv"
TMS_DEV_FILE = DEV_DIR / "TRRL0001_TMS.csv"
COLUMN_MAP_FILE = DEV_DIR / "column_map.json"

MERGED_DEV_FILE = DEV_DIR / "_merged.csv"
MERGED_PROD_CSV = PROD_DIR / "TRRL0001_merged.csv"
MERGED_PROD_XLSX = PROD_DIR / "TRRL0001_merged.xlsx"
MERGED_PROD_ODS = PROD_DIR / "TRRL0001_merged.ods"
MISSING_SHIPMENTS_FILE = PROD_DIR / "missing_shipments.txt"

JOIN_COLS = ["Shipment", "Container"]

# Explicit FIS -> TMS column map, covering every FIS column. Columns with
# no clear 1:1 TMS counterpart (e.g. Tmp. Sep., Handover Reference,
# Approval, ...) map to '' and are left un-renamed; they are skipped by
# compare_datasets since they won't have a match in TMS.
#
# Updated 2026-09-18 for the new abbreviated TMS export column names
# (Sh Ty, Ctpe, Rem, AutEx, Start-Loc, ... instead of the old spelled-out
# names). 'Dispatch Status' and 'Transportation Execution Status' have no
# confident new counterpart: the new TMS export dropped those and added
# 'TOC-Status'/'CWO-Status' instead, but which (if either) corresponds to
# which old field is unclear, so both are left unmapped ('') rather than
# guessed. Same for the new TMS-only columns 'Sel-Date', 'Sel-Time',
# 'CWO-Type', 'TOC', which have no FIS counterpart at all.
FIS_TMS_MAP = {
    'Dispatch Status ':               'TOC-Status',
    'Tmp. Sep.':                      '',
    'Transportation Execution Status':'CWO-Status',
    'Shipment Maintype':              'Sh Ty',
    'Shipment':                       'Shipment',
    'Sequence':                       'Seq',
    'MoT':                            'MoT',
    'CntrType':                       'Ctpe',
    'Weight':                         'Weight',
    'Unit':                           'Unit',
    'Remark':                         'Rem',
    'ProductIndicator':               'Prod Ind',
    'Temp[C]':                        'Temp \xa9',
    'DG':                             'DG',
    'DGAccept':                       'DG Accept',
    'Overdimension':                  'OD',
    'US Flag':                        'US Flg',
    'P/R/I':                          'PRI',
    'for':                            '',
    'ExcludeAuto Dispatch':           'AutEx',
    'AutoDispatchStatus':             'AutSt',
    'Multistop':                      'MS',
    'ImportExport':                   'I/E',
    'Start Locode':                   'Start-Loc',
    'Start Place Matchcode':          'Start-MC',
    'Start Date':                     'Start-Date',
    'Start Time':                     'Start-Time',
    'PosDel Locode':                  'PosDel-Loc',
    'PosDel Place Matchcode':         'PosDel-MC',
    'PosDel Postcode':                'PosDel-Post',
    'PosDel City':                    'PosDel-City',
    'PosDel Date':                    'PosDel-Date',
    'PosDel Time Sign':               '',
    'PosDel Time':                    'PosDel-Time',
    'End Locode':                     'End-Loc',
    'End Place Matchcode':            'End-MC',
    'End Date':                       'End-Date',
    'End Time':                       'End-Time',
    'Container Number':               'Container',
    'Container Rel.':                 '',
    'HandoverReference':              '',
    'LastFree Date':                  'LastFree',
    'Number':                         '',
    'SubconName':                     '',
    'SubconSupp':                     '',
    'Status':                         '',
    'GroupingNo.':                    '',
    'Number.1':                       'CWO',
    'Subcon Matchcode.1':             'Vendor-MC',
    'Approval':                       '',
    'Announced Delay':                '',
}

# Order in which compared columns should appear in the merged output,
# following the order they're declared in FIS_TMS_MAP.
COLUMN_ORDER = list(dict.fromkeys(v for v in FIS_TMS_MAP.values() if v))

# Blank strings, dash placeholders ('-', '--', ...) and common N/A markers
NULL_LIKE_PATTERN = re.compile(r'^\s*$|^-+$|^(n/?a|null|none|nan)$', re.IGNORECASE)

# TMS matchcode columns concatenate the FIS place Name with the zero-padded
# (3-digit) place Supp code, e.g. Name='BARRIO', Supp='38' -> 'BARRIO038'.
MATCHCODE_PAIRS = [
    ('Start Place Name', 'Start Place Supp', 'Start Place Matchcode'),
    ('PosDel Place Name', 'PosDel Place Supp', 'PosDel Place Matchcode'),
    ('End Place Name', 'End Place Supp', 'End Place Matchcode'),
    ('SubconName.1', 'SubconSupp.1', 'Subcon Matchcode.1'),
]

# FIS date/time columns are exported as full timestamps: date columns pad
# the time with a dummy 00:00:00, and time columns pad the date with a
# dummy 1970-01-01. Trim each down to just the part it actually carries.
# Note 'PosDel Time Sign' is not a timestamp (it holds a comparison
# operator like '<' or '='), so it's excluded from FIS_TIME_COLS.
FIS_DATE_COLS = ['Start Date', 'PosDel Date', 'End Date', 'LastFree Date']
FIS_TIME_COLS = ['Start Time', 'PosDel Time', 'End Time']


def create_dirs():
    for d in (RAW_DIR, DEV_DIR, PROD_DIR):
        d.mkdir(parents=True, exist_ok=True)


def build_fis_csv_from_xls():
    """Concatenate every FIS 'Transportation Request List' .xls export in
    FIS_RAW_XLS_DIR into a single FIS.csv, overwriting the raw csv.

    Each export has a 2-row header (group row + field row) above the data,
    plus a leading UI-only 'Status' column with no data-model equivalent.
    Rather than re-deriving flattened names from the group/field rows
    (unreliable - some groups are folded into the flat name, e.g. 'Start',
    others are dropped, e.g. 'Preadvice'), the data columns are matched
    positionally to FIS_COLUMNS, whose order matches this report layout.
    """
    xls_files = sorted(FIS_RAW_XLS_DIR.glob("*.xls"))
    frames = []
    for f in xls_files:
        raw = pd.read_excel(f, sheet_name="Transportation Request List", header=None, dtype=str)
        data = raw.iloc[3:, 1:]  # drop header rows (0-2) and the leading Status column
        data.columns = FIS_COLUMNS
        frames.append(data)
        print(f"[INFO] {f.name}: {len(data)} rows")

    df_fis = pd.concat(frames, ignore_index=True)
    df_fis.to_csv(FIS_RAW_CSV, index=False)
    print(f"[INFO] Concatenated {len(xls_files)} FIS exports ({len(df_fis)} rows total) -> {FIS_RAW_CSV}")
    return df_fis


def load_raw_datasets():
    df_fis = pd.read_csv(FIS_RAW_CSV, dtype=str)
    df_tms = pd.read_csv(TMS_RAW_FILE, dtype=str, encoding="utf-8-sig")

    print(f"[INFO] FIS columns: {list(df_fis.columns)}")
    print(f"[INFO] TMS columns: {list(df_tms.columns)}")

    return df_fis, df_tms


def clean_whitespace(df_fis, df_tms):
    """Collapse repeated internal whitespace and strip leading/trailing
    whitespace in all string columns, e.g. FIS 'Container Number' values
    like 'HLBU  1033723' (double space) vs TMS 'HLXU 8084231' (single
    space) would otherwise never match on join."""
    df_fis, df_tms = df_fis.copy(), df_tms.copy()

    for name, df in (('FIS', df_fis), ('TMS', df_tms)):
        for col in df.columns:
            if df[col].dtype != object:
                continue
            cleaned = df[col].str.replace(r'\s+', ' ', regex=True).str.strip()
            n_changed = int((cleaned != df[col]).fillna(False).sum())
            if n_changed:
                print(f"[INFO] {name} '{col}': normalized whitespace in {n_changed} values")
            df[col] = cleaned

    return df_fis, df_tms


def truncate_fis_datetimes(df_fis):
    """Trim FIS date columns down to 'YYYY-MM-DD' and time columns down to
    'HH:MM', dropping the dummy date/time part each carries (see
    FIS_DATE_COLS / FIS_TIME_COLS)."""
    df_fis = df_fis.copy()

    for col in FIS_DATE_COLS:
        df_fis[col] = df_fis[col].str.split(' ').str[0]

    for col in FIS_TIME_COLS:
        df_fis[col] = df_fis[col].str.split(' ').str[1].str[:5]

    return df_fis


def build_fis_matchcodes(df_fis):
    """Build FIS matchcode columns matching the TMS convention (Name + Supp
    zero-padded to 3 digits, no separator), replacing the separate Name/Supp
    columns so they can be compared against the TMS matchcode columns."""
    df_fis = df_fis.copy()

    for name_col, supp_col, matchcode_col in MATCHCODE_PAIRS:
        df_fis[matchcode_col] = df_fis[name_col] + df_fis[supp_col].str.zfill(3)
        print(f"[INFO] Built '{matchcode_col}' = '{name_col}' + '{supp_col}'.zfill(3)")

    name_supp_cols = [c for pair in MATCHCODE_PAIRS for c in pair[:2]]
    return df_fis.drop(columns=name_supp_cols)


def rename_columns(df_fis, df_tms):
    df_fis, df_tms = df_fis.copy(), df_tms.copy()

    # Strip trailing spaces from column names
    df_fis.columns = df_fis.columns.str.strip()
    df_tms.columns = df_tms.columns.str.strip()

    # Rename FIS columns to match TMS, skipping blank (unmapped) entries
    active_renames = {fis_col: tms_col for fis_col, tms_col in FIS_TMS_MAP.items() if tms_col}
    df_fis = df_fis.rename(columns=active_renames)

    print(f"[INFO] FIS columns (renamed): {list(df_fis.columns)}")
    print(f"[INFO] TMS columns: {list(df_tms.columns)}")

    return df_fis, df_tms


def check_columns_match(df_fis, df_tms):
    """Report columns that only exist in one dataset. These are expected here
    (FIS and TMS don't have a fully overlapping schema for TRRL0001) and will
    simply be skipped by compare_datasets, so this is informational only."""
    only_fis = sorted(set(df_fis.columns) - set(df_tms.columns))
    only_tms = sorted(set(df_tms.columns) - set(df_fis.columns))
    if only_fis:
        print(f"[INFO] Only in FIS (will be skipped in comparison): {only_fis}")
    if only_tms:
        print(f"[INFO] Only in TMS (will be skipped in comparison): {only_tms}")


def save_column_map(df_fis, df_tms):
    """Persist the FIS->TMS rename map plus which (renamed) columns are
    shared, FIS-only or TMS-only, for traceability of the comparison."""
    only_fis = sorted(set(df_fis.columns) - set(df_tms.columns))
    only_tms = sorted(set(df_tms.columns) - set(df_fis.columns))
    compared = sorted(set(df_fis.columns) & set(df_tms.columns))

    column_map = {
        "fis_tms_map": FIS_TMS_MAP,
        "compared_columns": compared,
        "only_in_fis": only_fis,
        "only_in_tms": only_tms,
    }

    with open(COLUMN_MAP_FILE, "w", encoding="utf-8") as f:
        json.dump(column_map, f, indent=2, ensure_ascii=False)

    print(f"[INFO] Column map saved -> {COLUMN_MAP_FILE}")


def clean_nulls(df_fis, df_tms):
    """Normalize null-like placeholder values (blank strings, '--', 'N/A', ...) to real NaN."""
    df_fis, df_tms = df_fis.copy(), df_tms.copy()

    for name, df in (('FIS', df_fis), ('TMS', df_tms)):
        for col in df.columns:
            if df[col].dtype != object:
                continue
            is_null_like = df[col].str.strip().str.match(NULL_LIKE_PATTERN).fillna(False).infer_objects(copy=False)
            mask = df[col].notna() & is_null_like
            n_found = int(mask.sum())
            if n_found:
                print(f"[INFO] {name} '{col}': normalized {n_found} null-like values to NaN")
            df.loc[mask, col] = pd.NA

    return df_fis, df_tms


def save_dev_datasets(df_fis, df_tms):
    df_fis.to_csv(FIS_DEV_FILE, index=False)
    df_tms.to_csv(TMS_DEV_FILE, index=False)
    print(f"[INFO] FIS dev dataset saved -> {FIS_DEV_FILE}")
    print(f"[INFO] TMS dev dataset saved -> {TMS_DEV_FILE}")


def write_formatted_xlsx(df, path):
    """Write df to xlsx with a dark-blue/bold/white header and green/red
    fill on '_match' and 'matched_in_*' columns depending on their Y/N value."""
    match_cols = [col for col in df.columns if col.endswith('_match') or col.startswith('matched_in_')]
    green_fill = PatternFill(start_color=MATCH_Y_BG, end_color=MATCH_Y_BG, fill_type="solid")
    red_fill = PatternFill(start_color=MATCH_N_BG, end_color=MATCH_N_BG, fill_type="solid")

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="merged")
        sheet = writer.sheets["merged"]

        for col_idx, col in enumerate(df.columns, start=1):
            width = min(max(len(str(col)), df[col].astype(str).str.len().max() if len(df) else 0) + 2, 40)
            sheet.column_dimensions[get_column_letter(col_idx)].width = width

        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name=FONT_NAME)

        for cell in sheet[1]:
            cell.font = Font(name=FONT_NAME, bold=True, size=HEADER_FONT_SIZE, color="FFFFFF")
            cell.fill = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type="solid")

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        last_row = len(df) + 1
        for col in match_cols:
            col_letter = get_column_letter(df.columns.get_loc(col) + 1)
            cell_range = f"{col_letter}2:{col_letter}{last_row}"
            sheet.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"Y"'], fill=green_fill))
            sheet.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"N"'], fill=red_fill))


def write_formatted_ods(df, path):
    """Write df to ods with a dark-blue/bold/white header and green/red
    fill on '_match' and 'matched_in_*' columns depending on their Y/N value."""
    match_cols = set(col for col in df.columns if col.endswith('_match') or col.startswith('matched_in_'))

    doc = OpenDocumentSpreadsheet()

    header_style = Style(name="Header", family="table-cell")
    header_style.addElement(TableCellProperties(backgroundcolor=f"#{HEADER_BG}"))
    header_style.addElement(TextProperties(
        color="#FFFFFF", fontweight="bold", fontsize=f"{HEADER_FONT_SIZE}pt", fontfamily=FONT_NAME,
    ))
    doc.automaticstyles.addElement(header_style)

    body_style = Style(name="Body", family="table-cell")
    body_style.addElement(TextProperties(fontfamily=FONT_NAME))
    doc.automaticstyles.addElement(body_style)

    match_y_style = Style(name="MatchY", family="table-cell")
    match_y_style.addElement(TableCellProperties(backgroundcolor=f"#{MATCH_Y_BG}"))
    match_y_style.addElement(TextProperties(fontfamily=FONT_NAME))
    doc.automaticstyles.addElement(match_y_style)

    match_n_style = Style(name="MatchN", family="table-cell")
    match_n_style.addElement(TableCellProperties(backgroundcolor=f"#{MATCH_N_BG}"))
    match_n_style.addElement(TextProperties(fontfamily=FONT_NAME))
    doc.automaticstyles.addElement(match_n_style)

    table = Table(name="merged")

    header_row = TableRow()
    for col in df.columns:
        cell = TableCell(stylename=header_style, valuetype="string")
        cell.addElement(P(text=str(col)))
        header_row.addElement(cell)
    table.addElement(header_row)

    for _, row in df.iterrows():
        tr = TableRow()
        for col in df.columns:
            value = row[col]
            text = '' if pd.isna(value) else str(value)

            style = body_style
            if col in match_cols:
                style = match_y_style if text == 'Y' else match_n_style if text == 'N' else body_style

            cell = TableCell(stylename=style, valuetype="string")
            cell.addElement(P(text=text))
            tr.addElement(cell)
        table.addElement(tr)

    doc.spreadsheet.addElement(table)
    doc.save(str(path))


def save_missing_shipments(df):
    """Save a SOQL query text file listing every unique Shipment number with
    no TMS match, for lookup against Transport_Order_Container__c."""
    missing = sorted(df.loc[df['matched_in_tms'] == 'N', 'Shipment'].dropna().unique())
    shipment_list = ", ".join(f"'{s}'" for s in missing)
    query = (
        "SELECT  Id, Transport_Order__r.Customer_Order__r.Type__c\n"
        "FROM Transport_Order_Container__c\n"
        f"WHERE Transport_Order__r.Customer_Order__r.Type__c IN ({shipment_list})\n"
    )
    MISSING_SHIPMENTS_FILE.write_text(query, encoding="utf-8")
    print(f"[INFO] Missing shipments query saved -> {MISSING_SHIPMENTS_FILE} ({len(missing)} shipments)")


def publish_merged_dataset():
    """Copy the dev merged csv to prod and render formatted xlsx/ods from it."""
    shutil.copyfile(MERGED_DEV_FILE, MERGED_PROD_CSV)
    print(f"[INFO] Merged dataset published -> {MERGED_PROD_CSV}")

    df = pd.read_csv(MERGED_PROD_CSV, dtype=str)

    write_formatted_xlsx(df, MERGED_PROD_XLSX)
    print(f"[INFO] Merged dataset published -> {MERGED_PROD_XLSX}")

    write_formatted_ods(df, MERGED_PROD_ODS)
    print(f"[INFO] Merged dataset published -> {MERGED_PROD_ODS}")

    save_missing_shipments(df)


def read():
    """Load the raw FIS/TMS datasets, concatenating the FIS .xls exports first."""
    create_dirs()
    build_fis_csv_from_xls()
    return load_raw_datasets()


def process(df_fis, df_tms):
    """Clean and align the raw datasets so they're ready to compare."""
    df_fis, df_tms = clean_whitespace(df_fis, df_tms)
    df_fis = truncate_fis_datetimes(df_fis)
    df_fis = build_fis_matchcodes(df_fis)
    df_fis, df_tms = rename_columns(df_fis, df_tms)
    check_columns_match(df_fis, df_tms)
    save_column_map(df_fis, df_tms)
    df_fis, df_tms = clean_nulls(df_fis, df_tms)

    save_dev_datasets(df_fis, df_tms)
    return df_fis, df_tms


def compare(df_fis, df_tms):
    """Merge/compare the processed datasets and publish the results."""
    summary = compare_datasets(
        df_fis, df_tms,
        join_cols=JOIN_COLS,
        output_dir=str(DEV_DIR),
        column_order=COLUMN_ORDER,
    )
    publish_merged_dataset()
    return summary


if __name__ == "__main__":
    df_fis, df_tms = read()
    df_fis, df_tms = process(df_fis, df_tms)
    summary = compare(df_fis, df_tms)
    print(summary)
