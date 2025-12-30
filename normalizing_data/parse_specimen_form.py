
import sys
import re
import pandas as pd
from datetime import datetime

# ---------------------------
# Helpers
# ---------------------------
def norm(s):
    """Strip whitespace and convert NaN to None."""
    if pd.isna(s):
        return None
    return str(s).strip()

def to_bool(v):
    """Map common yes/no-like values to boolean."""
    if v is None:
        return None
    s = str(v).strip().lower()
    if s in {"yes", "y", "true", "present", "1"}:
        return True
    if s in {"no", "n", "false", "absent", "0"}:
        return False
    return None

def to_num(v):
    """Extract the first numeric token, e.g., '~50' -> 50.0"""
    if v is None:
        return None
    m = re.search(r"([0-9]*\.?[0-9]+)", str(v))
    return float(m.group(1)) if m else None

def parse_date(v):
    """Parse a date-like value into a datetime.date, else None."""
    if v is None:
        return None
    # Try pandas first (handles many formats)
    try:
        return pd.to_datetime(v).date()
    except Exception:
        pass
    # Fallback formats
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(v), fmt).date()
        except Exception:
            continue
    return None

def to_yyyy_mm_dd(date_val):
    """Format date as YYYY-MM-DD string, None stays None."""
    if date_val is None or pd.isna(date_val):
        return None
    # date_val can be datetime.date or Timestamp
    try:
        return pd.to_datetime(date_val).strftime("%Y-%m-%d")
    except Exception:
        return None

def snake_case(name):
    """Convert a column name to lower_snake_case."""
    if name is None:
        return None
    s = str(name).strip()
    # Replace non-alphanumeric with spaces
    s = re.sub(r"[^0-9a-zA-Z]+", " ", s)
    # Collapse spaces and underscore
    s = "_".join(s.split())
    return s.lower()

def snake_case_df_columns(df):
    """Convert all columns of a DataFrame to lower_snake_case."""
    df = df.copy()
    df.columns = [snake_case(c) for c in df.columns]
    return df

# ---------------------------
# Main
# ---------------------------
def main():
    # CLI args
    INPUT = sys.argv[1] if len(sys.argv) > 1 else "30214.xlsx"
    OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "normalized.xlsx"

    # Read the Excel (first sheet)
    df = pd.read_excel(INPUT, engine="openpyxl")
    rows = df.values

    # ---- Specimen ID ----
    specimen_id = None
    # Prefer first column numeric token
    for val in df.iloc[:, 0]:
        if pd.notna(val) and re.fullmatch(r"\d+", str(val)):
            specimen_id = int(val)
            break
    # Fallback: scan entire sheet
    if specimen_id is None:
        for v in rows.flatten():
            if pd.notna(v) and re.fullmatch(r"\d+", str(v)):
                specimen_id = int(v)
                break

    # ---- Section parsing ----
    # Each section row: col A = section name; cols B.. = headers
    # Next row: values
    sections = {}
    for i in range(len(rows) - 1):
        sec_name = norm(rows[i][0])
        headers = [norm(x) for x in rows[i][1:]]
        if sec_name and any(h for h in headers):
            values = [norm(x) for x in rows[i + 1][1:]]
            mapping = {h: v for h, v in zip(headers, values) if h}
            sections[sec_name] = mapping

    # ---- Build normalized tables ----
    # From field sheet
    field = sections.get("From field sheet", {})

    # Capture duplicate 'Host tree' columns in order
    host_vals = []
    fs_idx = next((i for i in range(len(rows)) if norm(rows[i][0]) == "From field sheet"), None)
    if fs_idx is not None and fs_idx + 1 < len(rows):
        hdrs = [norm(x) for x in rows[fs_idx][1:]]
        vals = [norm(x) for x in rows[fs_idx + 1][1:]]
        host_vals = [vals[j] for j, h in enumerate(hdrs) if h == "Host tree"]

    # Specimens
    specimen_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Collector": field.get("Collector"),
        "SectionID": field.get("Section ID"),
        "HostTree1": host_vals[0] if host_vals else field.get("Host tree"),
        "HostTree2": host_vals[1] if len(host_vals) > 1 else None,
        "Coordinates": field.get("Coordinates"),
        "County": field.get("County"),
        "Site": field.get("Site"),
        "DateCollected": parse_date(field.get("Date")),
    }])

    # Photos
    photos_df = pd.DataFrame({
        "SpecimenID": specimen_id,
        "PhotoFile": [v for k, v in field.items() if k and k.lower().startswith("photo") and v]
    })

    # Pileus
    pileus = sections.get("Pileus", {})
    pileus_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Diameter_mm": to_num(pileus.get("Diameter (mm)")),
        "Thickness_mm": to_num(pileus.get("Thickness (mm)")),
        "CenterColor": pileus.get("Center color"),
        "MarginColor": pileus.get("Margin color"),
        "Shape": pileus.get("Shape"),
        "SurfaceTexture": pileus.get("Surface texture"),
        "Umbonate": to_bool(pileus.get("Umbonate?")),
        "StriationLength_mm": to_num(pileus.get("Striation length (mm)")),
        "Appendiculate": to_bool(pileus.get("Appendiculate?")),
        "AppendiculateTexture": pileus.get("Appendiculate texture"),
        "UniversalVeilPresent": to_bool(pileus.get("Universal veil")),
        "Staining": pileus.get("Staining"),
        "ContextColor": pileus.get("Context color"),
    }])

    # Universal veil on pileus
    uv_pileus = sections.get("Universal veil on pileus", {})
    uv_pileus_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Form": uv_pileus.get("Form"),
        "Color": uv_pileus.get("Color"),
        "Texture": uv_pileus.get("Texture"),
        "Attachment": uv_pileus.get("Attachment"),
        "Distribution": uv_pileus.get("Distribution"),
    }])

    # Lamellae
    lamellae = sections.get("Lamellae", {})
    lamellae_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Attachment": lamellae.get("Attachment"),
        "ColorInMasse": lamellae.get("Color in masse"),
        "Distribution": lamellae.get("Distribution"),
        "Staining": lamellae.get("Staining"),
        "OtherComment": lamellae.get("Other comment"),
    }])

    # Lamellulae
    lamellulae = sections.get("Lamellulae", {})
    lamellulae_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Form": lamellulae.get("Form"),
        "Amount": lamellulae.get("Amount"),
        "Distribution": lamellulae.get("Distribution"),
        "End": lamellulae.get("End"),
    }])

    # Stipe
    stipe = sections.get("Stipe", {})
    stipe_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Length_mm": to_num(stipe.get("Length (mm)")),
        "Width_mm": to_num(stipe.get("Width (mm)")),
        "Color": stipe.get("Color"),
        "Shape": stipe.get("Shape"),
        "Decoration": stipe.get("Decoration"),
        "Staining": stipe.get("Staining"),
        "AnnulusPresent": to_bool(stipe.get("Annulus")),
        "UniversalVeilPresent": to_bool(stipe.get("Universal veil")),
        "ContextType": stipe.get("Context type"),
        "ContextColor": stipe.get("Context color"),
        "ContextStain": stipe.get("Context stain"),
    }])

    # Annulus details
    ann = sections.get("Annulus", {})
    annulus_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Position": ann.get("Position"),
        "Form": ann.get("Form"),
        "Color": ann.get("Color"),
        "Staining": ann.get("Staining"),
        "RemainingPercent": to_num(ann.get("Remaining %")),
        "Condition": ann.get("Condition"),
    }])

    # Universal veil on stipe base
    uv_stipe_base = sections.get("Universal veil on stipe base", {})
    uv_stipe_base_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Type": uv_stipe_base.get("Type"),
        "Texture": uv_stipe_base.get("Texture"),
        "Color": uv_stipe_base.get("Color"),
        "Layered": to_bool(uv_stipe_base.get("Layered?")),
        "ToughOrFlimsy": uv_stipe_base.get("Tough/flimsy?"),
    }])

    # Basal bulb
    bulb = sections.get("Basal bulb", {})
    basal_bulb_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Length_mm": to_num(bulb.get("Length (mm)")),
        "Width_mm": to_num(bulb.get("Width (mm)")),
        "Shape": bulb.get("Shape"),
    }])

    # ---------------------------
    # Format dates and snake_case columns
    # ---------------------------
    # 1) Format DateCollected to YYYY-MM-DD where present.
    if "DateCollected" in specimen_df.columns:
        specimen_df["DateCollected"] = specimen_df["DateCollected"].apply(to_yyyy_mm_dd)

    # 2) Convert all DataFrame column names to lower_snake_case.
    specimen_df          = snake_case_df_columns(specimen_df)
    photos_df            = snake_case_df_columns(photos_df)
    pileus_df            = snake_case_df_columns(pileus_df)
    uv_pileus_df         = snake_case_df_columns(uv_pileus_df)
    lamellae_df          = snake_case_df_columns(lamellae_df)
    lamellulae_df        = snake_case_df_columns(lamellulae_df)
    stipe_df             = snake_case_df_columns(stipe_df)
    annulus_df           = snake_case_df_columns(annulus_df)
    uv_stipe_base_df     = snake_case_df_columns(uv_stipe_base_df)
    basal_bulb_df        = snake_case_df_columns(basal_bulb_df)

    # ---------------------------
    # Write output
    # ---------------------------
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as w:
        specimen_df.to_excel(w, index=False, sheet_name="specimens")
        photos_df.to_excel(w, index=False, sheet_name="photos")
        pileus_df.to_excel(w, index=False, sheet_name="pileus")
        uv_pileus_df.to_excel(w, index=False, sheet_name="universal_veil_pileus")
        lamellae_df.to_excel(w, index=False, sheet_name="lamellae")
        lamellulae_df.to_excel(w, index=False, sheet_name="lamellulae")
        stipe_df.to_excel(w, index=False, sheet_name="stipe")
        annulus_df.to_excel(w, index=False, sheet_name="annulus")
        uv_stipe_base_df.to_excel(w, index=False, sheet_name="universal_veil_stipe_base")
        basal_bulb_df.to_excel(w, index=False, sheet_name="basal_bulb")

    print(f"✅ Parsed {INPUT} → {OUTPUT}")

if __name__ == "__main__":
    main()
