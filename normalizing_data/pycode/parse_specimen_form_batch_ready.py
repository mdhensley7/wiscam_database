

import pandas as pd
import re
from datetime import datetime
import normalizing_data.pycode.parse_specimen_form as psf

# Reuse helper functions from parse_specimen_form
norm = psf.norm
to_bool = psf.to_bool
to_num = psf.to_num
parse_date = psf.parse_date
to_yyyy_mm_dd = psf.to_yyyy_mm_dd
snake_case_df_columns = psf.snake_case_df_columns


def parse_sheet(df: pd.DataFrame, specimen_id: str = None):
    """
    Parse a single 30214-style sheet DataFrame and return a dict of DataFrames:
      {
        'specimens': df1,
        'photos': df2,
        'pileus': df3,
        'universal_veil_pileus': df4,
        'lamellae': df5,
        'lamellulae': df6,
        'stipe': df7,
        'annulus': df8,
        'universal_veil_stipe_base': df9,
        'basal_bulb': df10,
      }
    If specimen_id is None, we attempt to derive it from the first numeric token
    found in column A or the sheet matrix.
    """
    rows = df.values

    # ---- specimen_id ----
    if specimen_id is None:
        specimen_id_local = None
        # Prefer first column numeric token
        for val in df.iloc[:, 0]:
            if pd.notna(val) and re.fullmatch(r"\d+", str(val)):
                specimen_id_local = int(val)
                break
        # Fallback: scan entire sheet
        if specimen_id_local is None:
            for v in rows.flatten():
                if pd.notna(v) and re.fullmatch(r"\d+", str(v)):
                    specimen_id_local = int(v)
                    break
        specimen_id = specimen_id_local

    # ---- sections ----
    sections = {}
    for i in range(len(rows) - 1):
        sec_name = norm(rows[i][0])
        headers = [norm(x) for x in rows[i][1:]]
        if sec_name and any(h for h in headers):
            values = [norm(x) for x in rows[i + 1][1:]]
            mapping = {h: v for h, v in zip(headers, values) if h}
            sections[sec_name] = mapping

    # From field sheet
    field = sections.get("From field sheet", {})

    # Capture duplicate 'Host tree' columns in order
    host_vals = []
    fs_idx = next((i for i in range(len(rows)) if norm(rows[i][0]) == "From field sheet"), None)
    if fs_idx is not None and fs_idx + 1 < len(rows):
        hdrs = [norm(x) for x in rows[fs_idx][1:]]
        vals = [norm(x) for x in rows[fs_idx + 1][1:]]
        host_vals = [vals[j] for j, h in enumerate(hdrs) if h == "Host tree"]

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

    photos_df = pd.DataFrame({
        "SpecimenID": specimen_id,
        "PhotoFile": [v for k, v in field.items() if k and k.lower().startswith("photo") and v]
    })

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

    uv_pileus = sections.get("Universal veil on pileus", {})
    uv_pileus_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Form": uv_pileus.get("Form"),
        "Color": uv_pileus.get("Color"),
        "Texture": uv_pileus.get("Texture"),
        "Attachment": uv_pileus.get("Attachment"),
        "Distribution": uv_pileus.get("Distribution"),
    }])

    lamellae = sections.get("Lamellae", {})
    lamellae_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Attachment": lamellae.get("Attachment"),
        "ColorInMasse": lamellae.get("Color in masse"),
        "Distribution": lamellae.get("Distribution"),
        "Staining": lamellae.get("Staining"),
        "OtherComment": lamellae.get("Other comment"),
    }])

    lamellulae = sections.get("Lamellulae", {})
    lamellulae_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Form": lamellulae.get("Form"),
        "Amount": lamellulae.get("Amount"),
        "Distribution": lamellulae.get("Distribution"),
        "End": lamellulae.get("End"),
    }])

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

    uv_stipe_base = sections.get("Universal veil on stipe base", {})
    uv_stipe_base_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Type": uv_stipe_base.get("Type"),
        "Texture": uv_stipe_base.get("Texture"),
        "Color": uv_stipe_base.get("Color"),
        "Layered": to_bool(uv_stipe_base.get("Layered?")),
        "ToughOrFlimsy": uv_stipe_base.get("Tough/flimsy?"),
    }])

    bulb = sections.get("Basal bulb", {})
    basal_bulb_df = pd.DataFrame([{
        "SpecimenID": specimen_id,
        "Length_mm": to_num(bulb.get("Length (mm)")),
        "Width_mm": to_num(bulb.get("Width (mm)")),
        "Shape": bulb.get("Shape"),
    }])

    # Format date and snake_case
    if "DateCollected" in specimen_df.columns:
        specimen_df["DateCollected"] = specimen_df["DateCollected"].apply(to_yyyy_mm_dd)

    specimen_df = snake_case_df_columns(specimen_df)
    photos_df = snake_case_df_columns(photos_df)
    pileus_df = snake_case_df_columns(pileus_df)
    uv_pileus_df = snake_case_df_columns(uv_pileus_df)
    lamellae_df = snake_case_df_columns(lamellae_df)
    lamellulae_df = snake_case_df_columns(lamellulae_df)
    stipe_df = snake_case_df_columns(stipe_df)
    annulus_df = snake_case_df_columns(annulus_df)
    uv_stipe_base_df = snake_case_df_columns(uv_stipe_base_df)
    basal_bulb_df = snake_case_df_columns(basal_bulb_df)

    return {
        'specimens': specimen_df,
        'photos': photos_df,
        'pileus': pileus_df,
        'universal_veil_pileus': uv_pileus_df,
        'lamellae': lamellae_df,
        'lamellulae': lamellulae_df,
        'stipe': stipe_df,
        'annulus': annulus_df,
        'universal_veil_stipe_base': uv_stipe_base_df,
        'basal_bulb': basal_bulb_df,
    }


def main():
    # Behave like the original main, but reuse parse_sheet
    import sys
    INPUT = sys.argv[1] if len(sys.argv) > 1 else "30214.xlsx"
    OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "normalized.xlsx"

    df = pd.read_excel(INPUT, engine="openpyxl")
    tables = parse_sheet(df)

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as w:
        for name, tdf in tables.items():
            tdf.to_excel(w, index=False, sheet_name=name)
    print(f"✅ Parsed {INPUT} → {OUTPUT}")

if __name__ == '__main__':
    main()
