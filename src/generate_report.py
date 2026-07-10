import os
import re

import pandas as pd


def _safe_filename(col):
    return re.sub(r'[\\/:*?"<>|]', '_', col)


def _format_combination_lines(combinations_path, top_n):
    combos = pd.read_csv(combinations_path)
    val_col1, val_col2 = combos.columns[0], combos.columns[1]
    label1 = val_col1.rsplit('_', 1)[-1].upper()
    label2 = val_col2.rsplit('_', 1)[-1].upper()

    def _display(val):
        return 'null' if pd.isna(val) else str(val).strip()

    lines = []
    for i, (_, row) in enumerate(combos.head(top_n).iterrows(), start=1):
        v1 = _display(row[val_col1])
        v2 = _display(row[val_col2])
        count = int(row['count'])
        pct = row['%_of_total']
        lines.append(f"{i}. {v1} ({label1}) and {v2} ({label2}): {pct}% ({count:,} instances).")

    if len(combos) > top_n:
        lines.append(f"{top_n + 1}. ...")

    return lines


def generate_report(dataset_dir, output_path=None, top_n=10):
    summary_path = os.path.join(dataset_dir, '_summary.csv')
    discrepancies_dir = os.path.join(dataset_dir, 'discrepancies')

    summary_df = pd.read_csv(summary_path).sort_values('column')

    title = os.path.basename(os.path.normpath(dataset_dir))
    if title in ('dev', 'prod'):
        title = os.path.basename(os.path.dirname(os.path.normpath(dataset_dir)))

    lines = [f"# {title} Discrepancy Report", '']

    for _, row in summary_df.iterrows():
        col = row['column']
        match_rate = row['match_rate_%']

        if row['discrepancies'] == 0:
            continue

        lines.append(f"### {col}")
        lines.append('')
        lines.append(f"{match_rate}%. _TODO: describe main discrepancy reason._")
        lines.append('')

        safe_col = _safe_filename(col)
        combinations_path = os.path.join(discrepancies_dir, f'{safe_col}_combinations.csv')
        if os.path.exists(combinations_path):
            lines.extend(_format_combination_lines(combinations_path, top_n))
        else:
            lines.append(f"_Missing combinations file: {combinations_path}_")

        lines.append('')

    report = '\n'.join(lines).rstrip() + '\n'

    output_path = output_path or os.path.join(dataset_dir, 'report.md')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(report)

    print(f"[INFO] Report saved -> {output_path}")
    return output_path


if __name__ == "__main__":

    REPORTS_LIST = ["TDWC0101"]

    BASE_DIR = "data"
    TOP_N = 10

    for name in REPORTS_LIST:
        for env in ("dev", "prod"):
            dataset_dir = os.path.join(BASE_DIR, name, env)
            if not os.path.exists(os.path.join(dataset_dir, '_summary.csv')):
                print(f"[SKIP] {dataset_dir} has no _summary.csv")
                continue
            generate_report(dataset_dir, top_n=TOP_N)
