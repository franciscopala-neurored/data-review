"""
generates the data directory and the data files for the TDWC0101 dataset.
"""

import re

import pandas as pd
from pathlib import Path
from compare_datasets import compare_datasets


DATA_DIR = Path("./data/TDWC0101")
RAW_DIR = DATA_DIR / "raw"
DEV_DIR = DATA_DIR / "dev"
PROD_DIR = DATA_DIR / "prod"

FIS_RAW_FILE = RAW_DIR / "TDWC0101_FIS.csv"
TMS_RAW_FILE = RAW_DIR / "TDWC0101_TMS.csv"
    
JOIN_COLS = ["Number", "Cont. Number"]

FIS_RENAME_MAP = {
    'Locode':            'Start Locode',
    'Start':             'Start Location',
    'Locode.1':          'Pos/Del Locode',
    'Pos/Del':           'Pos/Del Location',
    'Locode.2':          'End Locode',
    'End':               'End Location',
    'Date':              'Pos/Del Date',
    'Time.1':            'Pos/Del Time',
    'Locode.3':          'Multistop Locode',
    'Multistop':         'Multistop Location',
    'OW Surcharge':      'Overweight Surcharge',
    'MautToll Surcharge':'Maut/Toll Surcharge',
}

# Blank strings, dash placeholders ('-', '--', ...) and common N/A markers
NULL_LIKE_PATTERN = re.compile(r'^\s*$|^-+$|^(n/?a|null|none|nan)$', re.IGNORECASE)


def create_dirs():
    for d in (RAW_DIR, DEV_DIR, PROD_DIR):
        d.mkdir(parents=True, exist_ok=True)


def load_raw_datasets():
    df_fis = pd.read_excel(FIS_RAW_FILE, engine="odf", sheet_name="data", header=1, dtype=str)
    df_tms = pd.read_csv(TMS_RAW_FILE, dtype=str)

    print(f"[INFO] FIS columns: {list(df_fis.columns)}")
    print(f"[INFO] TMS columns: {list(df_tms.columns)}")

    return df_fis, df_tms


def rename_columns(df_fis, df_tms):
    df_fis, df_tms = df_fis.copy(), df_tms.copy()

    # Strip trailing spaces from TMS column names
    df_tms.columns = df_tms.columns.str.strip()

    # Rename FIS columns to match TMS
    df_fis = df_fis.rename(columns=FIS_RENAME_MAP).drop(columns=['Unnamed: 19'], errors='ignore')

    print(f"[INFO] FIS columns: {list(df_fis.columns)}")
    print(f"[INFO] TMS columns: {list(df_tms.columns)}")

    return df_fis, df_tms


def check_columns_match(df_fis, df_tms):
    only_fis = set(df_fis.columns) - set(df_tms.columns)
    only_tms = set(df_tms.columns) - set(df_fis.columns)
    if only_fis or only_tms:
        raise ValueError(
            f"Column mismatch between FIS and TMS after renaming. "
            f"Only in FIS: {sorted(only_fis)}. Only in TMS: {sorted(only_tms)}."
        )


def clean_nulls(df_fis, df_tms):
    """Normalize null-like placeholder values (blank strings, '--', 'N/A', ...) to real NaN."""
    df_fis, df_tms = df_fis.copy(), df_tms.copy()

    for name, df in (('FIS', df_fis), ('TMS', df_tms)):
        for col in df.columns:
            if df[col].dtype != object:
                continue
            is_null_like = df[col].str.strip().str.match(NULL_LIKE_PATTERN).fillna(False).astype(bool)
            mask = df[col].notna() & is_null_like
            n_found = int(mask.sum())
            if n_found:
                print(f"[INFO] {name} '{col}': normalized {n_found} null-like values to NaN")
            df.loc[mask, col] = pd.NA

    return df_fis, df_tms


# TODO: publish the cleaned dataset to DATA_DIR / dev / datasets


if __name__ == "__main__":
    create_dirs()

    df_fis, df_tms = load_raw_datasets()
    df_fis, df_tms = rename_columns(df_fis, df_tms)
    check_columns_match(df_fis, df_tms)
    df_fis, df_tms = clean_nulls(df_fis, df_tms)

    summary = compare_datasets(
        df_fis, df_tms,
        join_cols=JOIN_COLS,
        output_dir=str(DEV_DIR),
    )
    print(summary)
