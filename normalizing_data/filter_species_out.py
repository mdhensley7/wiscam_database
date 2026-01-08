
import pandas as pd
from pathlib import Path

# ====== CONFIG ======
INPUT_XLSX = "specimens_long.xlsx"            # change if your file has a different name/path
OUTPUT_XLSX = "specimens_long_filtered.xlsx"  # new file will be written here

# Treat these as "missing" values in the species column (case-insensitive, trimmed)
EMPTY_SPECIES = {"", "na", "n/a", "null", "none"}

def is_valid_species(val) -> bool:
    """Return True iff species is a meaningful (non-empty, non-placeholder) string."""
    if pd.isna(val):
        return False
    s = str(val).strip()
    return (len(s) > 0) and (s.lower() not in EMPTY_SPECIES)

def filter_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Drop rows where species is missing/placeholder; if no species column, return as-is."""
    if "species" not in df.columns:
        # Some sheets might not have 'species'; we leave them untouched
        return df
    # Keep rows with valid species
    mask = df["species"].apply(is_valid_species)
    return df.loc[mask].copy()

def main():
    src_path = Path(INPUT_XLSX)
    if not src_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {src_path.resolve()}")

    # Read original workbook
    xl = pd.ExcelFile(src_path, engine="openpyxl")

    # Process each sheet, dropping completely empty rows first, then filtering by species
    filtered = {}
    summary = []

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, dtype=str)  # load as strings to keep values consistent
        df = df.dropna(how="all")             # drop rows that are entirely empty
        df_filtered = filter_sheet(df)

        filtered[sheet_name] = df_filtered

        summary.append({
            "sheet": sheet_name,
            "original_rows": len(df),
            "filtered_rows": len(df_filtered)
        })

    # Write the filtered workbook
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, df in filtered.items():
            df.to_excel(writer, sheet_name=name, index=False)

    # Print a quick summary to the console
    print("Done. Wrote:", OUTPUT_XLSX)
    print("Row counts by sheet:")
    for item in summary:
        print(f"  - {item['sheet']}: {item['original_rows']} → {item['filtered_rows']}")

if __name__ == "__main__":
    main()
