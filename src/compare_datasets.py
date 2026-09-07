"""

"""

import pandas as pd
import os
import re
from pathlib import Path


def _safe_filename(col):
    return re.sub(r'[\\/:*?"<>|]', '_', col)


def _norm(series):
    if series.dtype == object:
        return series.where(series.isna(), series.astype(str).str.strip().str.upper())
    return series


def compare_datasets(df1, df2, join_cols, output_dir, df1_name='fis', df2_name='tms'):
    """
    the function sequentially compares the columns of two datasets.
    the names must match.
    the merged dataset is saved in the output_dir, with a summary of the comparison.
    """
    matches_dir = os.path.join(output_dir, 'matches')
    discrepancies_dir = os.path.join(output_dir, 'discrepancies')
    os.makedirs(matches_dir, exist_ok=True)
    os.makedirs(discrepancies_dir, exist_ok=True)

    df1, df2 = df1.copy(), df2.copy()
    for col in join_cols:
        t1, t2 = df1[col].dtype, df2[col].dtype
        if t1 != t2:
            print(f"[WARN] '{col}' type mismatch: df1={t1}, df2={t2} — casting both to str")
            df1[col] = df1[col].astype(str)
            df2[col] = df2[col].astype(str)
        for df in (df1, df2):
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True).str.upper()

    all_cols_1 = set(df1.columns) - set(join_cols)
    all_cols_2 = set(df2.columns) - set(join_cols)
    common_cols = sorted(all_cols_1 & all_cols_2)
    only_in_1 = sorted(all_cols_1 - all_cols_2)
    only_in_2 = sorted(all_cols_2 - all_cols_1)

    print(f"[INFO] Columns in both (will compare): {common_cols}")
    if only_in_1:
        print(f"[SKIP] Only in {df1_name} (skipped): {only_in_1}")
    if only_in_2:
        print(f"[SKIP] Only in {df2_name} (skipped): {only_in_2}")
    print()

    merged = pd.merge(
        df1,
        df2,
        on=join_cols,
        how='inner',
        suffixes=(f'_{df1_name}', f'_{df2_name}')
    )

    summary_rows = []

    for col in common_cols:
        col1, col2 = f'{col}_{df1_name}', f'{col}_{df2_name}'

        norm1, norm2 = _norm(merged[col1]), _norm(merged[col2])
        disc_mask = ~(
            (norm1 == norm2) |
            (norm1.isna() & norm2.isna())
        )
        match_col = f'{col}_match'
        merged[match_col] = 'Y'
        merged.loc[disc_mask, match_col] = 'N'

        discrepancies = merged[disc_mask]
        matches = merged[~disc_mask]

        total = len(merged)
        n_disc = len(discrepancies)
        n_match = len(matches)
        match_rate = round(n_match / total, 4) if total > 0 else None
        disc_rate = round(n_disc / total, 4) if total > 0 else None

        notna_match_mask = (~disc_mask) & norm1.notna() & norm2.notna()
        n_notna_match = int(notna_match_mask.sum())
        notna_match_rate = round(n_notna_match / total * 100, 2) if total > 0 else None

        summary_rows.append({
            'column': col,
            'total_rows': total,
            'matching_rows': n_match,
            'discrepancies': n_disc,
            'match_rate_%': round(match_rate * 100, 2) if match_rate is not None else None,
            'discrepancy_rate_%': round(disc_rate * 100, 2) if disc_rate is not None else None,
            'notna_match_rows': n_notna_match,
            'notna_match_rate_%': notna_match_rate,
        })

        safe_col = _safe_filename(col)

        matches_path = os.path.join(matches_dir, f'{safe_col}_matches.csv')
        matches[join_cols + [col1, col2]].to_csv(matches_path, index=False)

        if n_disc > 0:
            disc_path = os.path.join(discrepancies_dir, f'{safe_col}.csv')
            discrepancies[join_cols + [col1, col2]].to_csv(disc_path, index=False)

            combinations = (
                discrepancies.groupby([col1, col2], dropna=False)
                .size()
                .reset_index(name='count')
                .sort_values('count', ascending=False)
            )
            combinations['%_of_total'] = round(combinations['count'] / total * 100, 2)
            combinations_path = os.path.join(discrepancies_dir, f'{safe_col}_combinations.csv')
            combinations.to_csv(combinations_path, index=False)

            print(f"[WARN] '{col}': {n_disc}/{total} mismatches -> {disc_path}")
        else:
            print(f"[OK]   '{col}': no discrepancies")

    ordered_cols = join_cols + sorted(c for c in merged.columns if c not in join_cols)
    merged = merged[ordered_cols]
    merged_path = os.path.join(output_dir, '_merged.csv')
    merged.to_csv(merged_path, index=False)
    print(f"\n[INFO] Merged dataset saved -> {merged_path}")

    summary_df = pd.DataFrame(summary_rows).sort_values('discrepancies', ascending=False)
    summary_path = os.path.join(output_dir, '_summary.csv')
    summary_df.to_csv(summary_path, index=False)
    print(f"[INFO] Summary saved -> {summary_path}")

    return summary_df


if __name__ == "__main__":
    pass
