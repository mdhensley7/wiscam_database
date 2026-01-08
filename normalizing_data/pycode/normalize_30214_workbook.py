
import argparse
import pandas as pd
import numpy as np
import re

num_re = re.compile(r"^-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?$")

def is_empty(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip())

def is_num_like(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        return True
    if isinstance(v, str):
        return num_re.match(v.strip()) is not None
    return False

def to_num(v):
    if isinstance(v, (int, float)):
        return float(v) if not pd.isna(v) else np.nan
    if isinstance(v, str):
        s = v.strip()
        if num_re.match(s):
            try:
                return float(s)
            except:
                return np.nan
    return np.nan

def snake_case(text: str) -> str:
    text = re.sub(r"[^0-9A-Za-z]+", "_", str(text).strip())
    text = re.sub(r"_+", "_", text)
    return text.strip("_").lower()

def normalize_sheet(df: pd.DataFrame, strict_two_col: bool = True):
    records = []
    rows, cols = df.shape
    for r in range(rows):
        label = None
        label_c = None
        value = None
        value_c = None
        if strict_two_col and cols >= 2:
            label = df.iat[r, 0]
            if isinstance(label, str) and label.strip():
                label_c = 0
                # take first non-empty value to the right (col1 then col2...)
                for c in range(1, cols):
                    v = df.iat[r, c]
                    if not is_empty(v):
                        value = v
                        value_c = c
                        break
        else:
            # fallback: first non-empty string in the row, then first non-empty value to the right
            for c in range(cols):
                v = df.iat[r, c]
                if isinstance(v, str) and v.strip():
                    label = v.strip(); label_c = c
                    break
            if label is not None:
                for c in range(label_c+1, cols):
                    w = df.iat[r, c]
                    if not is_empty(w):
                        value = w; value_c = c
                        break
        if label is None or value is None:
            continue
        records.append({
            'key': str(label).strip(),
            'value': str(value) if not isinstance(value, (int, float)) else str(value),
            'value_num': to_num(value),
            'row': r,
            'label_col': label_c,
            'value_col': value_c
        })
    return pd.DataFrame(records)

def compile_workbook(path: str, strict_two_col: bool = True, sanitize_keys: bool = True, last_only: bool = False):
    xl = pd.ExcelFile(path, engine='openpyxl')
    frames = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)
        norm = normalize_sheet(df, strict_two_col=strict_two_col)
        if norm.empty:
            continue
        norm['specimen_id'] = sheet
        frames.append(norm)
    if not frames:
        return pd.DataFrame(), pd.DataFrame()
    long_df = pd.concat(frames, ignore_index=True)

    # De-duplicate: keep last occurrence per specimen_id + key if requested
    if last_only:
        long_df['order'] = long_df.groupby('specimen_id').cumcount()
        long_df = long_df.sort_values(['specimen_id','key','order'])
        long_df = long_df.drop_duplicates(['specimen_id','key'], keep='last')
        long_df = long_df.drop(columns=['order'])

    # Wide pivot using numeric-first choice
    wide_df = long_df.copy()
    wide_df['value_choice'] = np.where(~wide_df['value_num'].isna(), wide_df['value_num'], wide_df['value'])
    if sanitize_keys:
        wide_df['key_sanitized'] = wide_df['key'].apply(snake_case)
        key_col = 'key_sanitized'
    else:
        key_col = 'key'
    wide_pivot = wide_df.pivot_table(index='specimen_id', columns=key_col, values='value_choice', aggfunc=lambda x: x.iloc[-1])
    wide_pivot = wide_pivot.reset_index()

    return long_df, wide_pivot

def main():
    import argparse
    p = argparse.ArgumentParser(description='Normalize 30214-style multi-sheet workbook (specimen_id = sheet title).')
    p.add_argument('workbook', help='Path to .xlsx workbook with >100 sheets')
    p.add_argument('-o','--out-prefix', default='specimens', help='Output file prefix')
    p.add_argument('--loose', action='store_true', help='Use loose row-wise parsing (not strictly two-column).')
    p.add_argument('--no-sanitize', action='store_true', help='Do not snake_case keys in wide output.')
    p.add_argument('--last-only', action='store_true', help='Keep only last occurrence per key within a sheet.')
    args = p.parse_args()

    strict_two_col = not args.loose
    sanitize_keys = not args.no_sanitize

    long_df, wide_df = compile_workbook(args.workbook, strict_two_col=strict_two_col, sanitize_keys=sanitize_keys, last_only=args.last_only)
    if long_df.empty:
        print('No data extracted; confirm the workbook matches the 30214 two-column format.')
        return

    # Write outputs
    long_csv = f"{args.out_prefix}_long.csv"
    long_xlsx = f"{args.out_prefix}_long.xlsx"
    wide_csv = f"{args.out_prefix}_wide.csv"
    wide_xlsx = f"{args.out_prefix}_wide.xlsx"

    long_df.to_csv(long_csv, index=False)
    wide_df.to_csv(wide_csv, index=False)
    with pd.ExcelWriter(long_xlsx, engine='openpyxl') as xw:
        long_df.to_excel(xw, index=False, sheet_name='long')
    with pd.ExcelWriter(wide_xlsx, engine='openpyxl') as xw:
        wide_df.to_excel(xw, index=False, sheet_name='wide')

    # Also emit a quick key coverage report (fill rate per key)
    coverage = (long_df.assign(non_empty=~long_df['value'].astype(str).str.strip().eq(''))
                .groupby('key')['non_empty'].mean().sort_values(ascending=False).rename('fill_rate').reset_index())
    coverage_csv = f"{args.out_prefix}_key_coverage.csv"
    coverage.to_csv(coverage_csv, index=False)

    print({'long_csv': long_csv, 'wide_csv': wide_csv, 'long_xlsx': long_xlsx, 'wide_xlsx': wide_xlsx, 'key_coverage_csv': coverage_csv})

if __name__ == '__main__':
    main()
