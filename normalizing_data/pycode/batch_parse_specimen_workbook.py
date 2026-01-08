
import argparse
import importlib
import sys
import pandas as pd
import numpy as np
import inspect

def snake_case(s: str) -> str:
    import re
    s = re.sub(r"[^0-9A-Za-z]+", "_", str(s).strip())
    s = re.sub(r"_+", "_", s)
    return s.strip('_').lower()

# ------------------ parser resolution ------------------

def load_module(module_name: str):
    try:
        return importlib.import_module(module_name)
    except Exception as e:
        print(f"ERROR: Could not import {module_name}. Ensure it is in the same folder or PYTHONPATH.")
        print(e)
        sys.exit(1)


def find_callable(mod, func_name: str = None, class_name: str = None, method_name: str = None):
    if class_name and method_name:
        cls = getattr(mod, class_name, None)
        if cls is None:
            raise RuntimeError(f"Class '{class_name}' not found in module.")
        inst = cls()
        func = getattr(inst, method_name, None)
        if not callable(func):
            raise RuntimeError(f"Method '{method_name}' not found or not callable on class '{class_name}'.")
        return func
    if func_name:
        func = getattr(mod, func_name, None)
        if not callable(func):
            raise RuntimeError(f"Function '{func_name}' not found or not callable in module.")
        return func
    candidate_names = [
        'parse_sheet', 'parse_specimen', 'parse', 'parse_form', 'parse_specimen_form',
        'parse_one', 'parse_workbook_sheet', 'extract_specimen', 'extract_sheet'
    ]
    for name in candidate_names:
        func = getattr(mod, name, None)
        if callable(func):
            return func
    for name, obj in inspect.getmembers(mod, predicate=callable):
        if any(tok in name.lower() for tok in ('parse','extract')):
            return obj
    func = getattr(mod, 'PARSER_FUNC', None)
    if callable(func):
        return func
    callables = [name for name, obj in inspect.getmembers(mod, predicate=callable)]
    raise RuntimeError("No suitable parser function found. Provide --func or --class/--method. "
                       f"Available callables: {callables}")


def call_parser(func, df, sheet_name: str):
    try:
        out = func(df, sheet_name)
    except TypeError:
        out = func(df)
    return out

# ------------------ wide pivot ------------------

def to_wide(long_df: pd.DataFrame, sanitize_keys: bool = True) -> pd.DataFrame:
    if 'value_num' in long_df.columns:
        long_df['value_choice'] = np.where(~long_df['value_num'].isna(), long_df['value_num'], long_df.get('value', np.nan))
    else:
        long_df['value_choice'] = long_df.get('value', np.nan)
    key_col = 'key'
    if sanitize_keys and 'key' in long_df.columns:
        long_df['key_sanitized'] = long_df['key'].apply(snake_case)
        key_col = 'key_sanitized'
    wide = (long_df
            .sort_values(['specimen_id', key_col])
            .pivot_table(index='specimen_id', columns=key_col, values='value_choice', aggfunc=lambda x: x.iloc[-1])
            .reset_index())
    return wide

# ------------------ main ------------------

def main():
    ap = argparse.ArgumentParser(description='Batch-run a specimen parser across all sheets in a workbook.')
    ap.add_argument('workbook', help='Path to the .xlsx workbook')
    ap.add_argument('-m','--module', default='parse_specimen_form_batch_ready', help='Parser module (default: parse_specimen_form_batch_ready)')
    ap.add_argument('-o','--out-prefix', default='specimens', help='Output file prefix')
    ap.add_argument('--func', help='Function name in the parser module to call per sheet (default: parse_sheet)')
    ap.add_argument('--class', dest='cls', help='Class name in the parser module (use with --method)')
    ap.add_argument('--method', help='Method name on the class (use with --class)')
    ap.add_argument('--no-sanitize', action='store_true', help='Do not snake_case keys in wide output')
    ap.add_argument('--last-only', action='store_true', help='If parser returns multiple rows per (specimen_id, key), keep only last occurrence')
    args = ap.parse_args()

    xl = pd.ExcelFile(args.workbook, engine='openpyxl')
    mod = load_module(args.module)
    parser_fn = find_callable(mod, func_name=args.func or 'parse_sheet', class_name=args.cls, method_name=args.method)

    combined_tables = {}  # name -> list of dfs
    long_frames = []

    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)
        try:
            parsed = call_parser(parser_fn, df, sheet)
        except Exception as e:
            print(f"WARN: Skipping sheet '{sheet}' due to parser error: {e}")
            continue

        # Accept either a DataFrame or a dict of DataFrames
        if isinstance(parsed, pd.DataFrame):
            parsed['specimen_id'] = sheet
            long_frames.append(parsed)
        elif isinstance(parsed, dict):
            for name, tdf in parsed.items():
                if tdf is None or tdf.empty:
                    continue
                tdf = tdf.copy()
                tdf['specimen_id'] = sheet
                combined_tables.setdefault(name, []).append(tdf)
            # If a 'specimens' table exists, also treat it as long frame
            if 'specimens' in parsed:
                sp = parsed['specimens'].copy()
                sp['specimen_id'] = sheet
                long_frames.append(sp)
        else:
            print(f"WARN: Parser output for sheet '{sheet}' is neither DataFrame nor dict; skipping.")
            continue

    if not long_frames and not combined_tables:
        print('No sheets parsed. Check parser signature or workbook format.')
        return

    # Combine long frames (specimens)
    long_df = pd.concat(long_frames, ignore_index=True) if long_frames else pd.DataFrame()

    # Optional: keep only last occurrence per (specimen_id, key) if such columns exist
    if args.last_only and {'specimen_id','key'}.issubset(long_df.columns):
        long_df['_order'] = long_df.groupby(['specimen_id']).cumcount()
        long_df = long_df.sort_values(['specimen_id','key','_order']).drop_duplicates(['specimen_id','key'], keep='last')
        long_df = long_df.drop(columns=['_order'])

    # Wide pivot only if long_df has 'key' and value/value_num columns; else skip
    wide_df = None
    if {'specimen_id','key'}.issubset(long_df.columns):
        wide_df = to_wide(long_df, sanitize_keys=not args.no_sanitize)

    # Write outputs
    long_csv = f"{args.out_prefix}_long.csv"
    long_xlsx = f"{args.out_prefix}_long.xlsx"
    long_df.to_csv(long_csv, index=False)
    with pd.ExcelWriter(long_xlsx, engine='openpyxl') as xw:
        long_df.to_excel(xw, index=False, sheet_name='specimens')
        # Also write combined tables if present
        for name, frames in combined_tables.items():
            full = pd.concat(frames, ignore_index=True)
            full.to_excel(xw, index=False, sheet_name=name[:31])

    if wide_df is not None:
        wide_csv = f"{args.out_prefix}_wide.csv"
        wide_xlsx = f"{args.out_prefix}_wide.xlsx"
        wide_df.to_csv(wide_csv, index=False)
        with pd.ExcelWriter(wide_xlsx, engine='openpyxl') as xw:
            wide_df.to_excel(xw, index=False, sheet_name='wide')
        print({'long_csv': long_csv, 'long_xlsx': long_xlsx, 'wide_csv': wide_csv, 'wide_xlsx': wide_xlsx})
    else:
        print({'long_csv': long_csv, 'long_xlsx': long_xlsx})

if __name__ == '__main__':
    main()
