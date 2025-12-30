# Normalizing Morphological Data
The first thing that needs to happen is to organize the data in order to for it to be easy to use when building the relational database

Since I set up my data in a poorly organized way, I had copilot help me take my data and organize it into a new sheet that is fully normalized

## Troubleshooting initial code

#### Here is the python code copilot gave me to organize a single specimen (30214)

```python
#Here are the needed packages
import sys, re
import pandas as pd
from datetime import datetime

INPUT = sys.argv[1] if len(sys.argv) > 1 else "30214_test.xlsx"
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "normalized.xlsx"

def norm(s):
    if pd.isna(s): return None
    return str(s).strip()

def to_bool(v):
    if v is None: return None
    s = str(v).strip().lower()
    if s in {"yes","y","true","present"}: return True
    if s in {"no","n","false","absent"}: return False
    return None

def to_num(v):
    if v is None: return None
    m = re.search(r"([0-9]*\\.?[0-9]+)", str(v))
    return float(m.group(1)) if m else None

def parse_date(v):
    if v is None: return None
    try:
        return pd.to_datetime(v).date()
    except Exception:
        for fmt in ("%m/%d/%Y","%Y-%m-%d","%d/%m/%Y"):
            try: return datetime.strptime(str(v), fmt).date()
            except Exception: pass
    return None

df = pd.read_excel(INPUT, engine="openpyxl")
rows = df.values

# ---- Specimen ID ----
specimen_id = None
for val in df.iloc[:,0]:
    if pd.notna(val) and re.fullmatch(r"\\d+", str(val)):
        specimen_id = int(val); break
if specimen_id is None:
    for v in rows.flatten():
        if pd.notna(v) and re.fullmatch(r"\\d+", str(v)):
            specimen_id = int(v); break

# ---- Section parsing ----
sections = {}
for i in range(len(rows)-1):
    sec_name = norm(rows[i][0])
    headers  = [norm(x) for x in rows[i][1:]]
    if sec_name and any(h for h in headers):
        values = [norm(x) for x in rows[i+1][1:]]
        mapping = {h:v for h,v in zip(headers, values) if h}
        sections[sec_name] = mapping

# ---- Build normalized tables ----
field = sections.get("From field sheet", {})
# capture duplicate "Host tree" columns
host_vals = []
# find the exact row to preserve order
fs_idx = next((i for i in range(len(rows)) if norm(rows[i][0])=="From field sheet"), None)
if fs_idx is not None:
    hdrs = [norm(x) for x in rows[fs_idx][1:]]
    vals = [norm(x) for x in rows[fs_idx+1][1:]]
    host_vals = [vals[j] for j,h in enumerate(hdrs) if h=="Host tree"]

specimen_df = pd.DataFrame([{
    "SpecimenID": specimen_id,
    "Collector": field.get("Collector"),
    "SectionID": field.get("Section ID"),
    "HostTree1": host_vals[0] if host_vals else field.get("Host tree"),
    "HostTree2": host_vals[1] if len(host_vals)>1 else None,
    "Coordinates": field.get("Coordinates"),
    "County": field.get("County"),
    "Site": field.get("Site"),
    "DateCollected": parse_date(field.get("Date")),
}])

photos_df = pd.DataFrame({
    "SpecimenID": specimen_id,
    "PhotoFile": [v for k,v in field.items() if k and k.lower().startswith("photo") and v]
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

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as w:
    specimen_df.to_excel(w, index=False, sheet_name="Specimens")
    photos_df.to_excel(w, index=False, sheet_name="Photos")
    pileus_df.to_excel(w, index=False, sheet_name="Pileus")
    uv_pileus_df.to_excel(w, index=False, sheet_name="UniversalVeil_Pileus")
    lamellae_df.to_excel(w, index=False, sheet_name="Lamellae")
    lamellulae_df.to_excel(w, index=False, sheet_name="Lamellulae")
    stipe_df.to_excel(w, index=False, sheet_name="Stipe")
    annulus_df.to_excel(w, index=False, sheet_name="Annulus")
    uv_stipe_base_df.to_excel(w, index=False, sheet_name="UniversalVeil_StipeBase")
    basal_bulb_df.to_excel(w, index=False, sheet_name="BasalBulb")

print(f"✅ Parsed {INPUT} → {OUTPUT}")
```

I saved this as /normalizing_data/parse_specimen_form.py

- Once saved, I changed my input file and output file to test the script would work again.
- It did not because I did not have the packages installed in the interpreter I am using
- Because of this, I have to create a virtual environment within my 'normalizing_data' folder which hosts the python script. The code for doing this is below

```zsh
cd /Users/hensley/Desktop/wiscam_database/normalizing_data

#check python is installed
python3 --version

#create a virtual environment within normalizing_data
python -m venv .venv

# Activate
source .venv/bin/activate
```

- A virtual environment avoids conflicts by setting up an environment within a project with which to download packages etc.

Once this is done the correct packages can be installed 

```zsh
python -m pip install --upgrade pip
python -m pip install pandas openpyxl
```

Now we can run the script with our new test .xlsx ('30181_test2.xlsx')

```zsh
cd /Users/hensley/Desktop/wiscam_database/normalizing_data
python parse_specimen_form.py "/Users/hensley/Desktop/wiscam_database/normalizing_data/30181_test2.xlsx" "/Users/hensley/Desktop/wiscam_database/normalizing_data/30181_test2_normalized.xlsx" #the full path is overboard since all should be in the same folder but it should work either way
```

Damn yeah that worked really well

### Analyzing test sheet
I need to figure out what needs to go and what isn't going to work very well because of how messy my data is set up

#### Here is list of changes to make:

1. specimenid is just going to be weird but we can manually input these because they will just be AmanitaBase numbers so shouldn't take too much time

2. I will need to go back and double check all section IDs because a lot of them are initially wrong

3. I will also need to manually scan host trees

4. I think I can remove the 'photos' sheet entirely so will delete corresponding code from the parser

5. Need to go through each of the sheets and add code that will fill in blanks correctly (ask copilot what the most appropriate input would be to make it easier to organize in a relational data base)
    - Some of the blanks just mean absent while some mean not applicable because you can't take a measurement of something absent

    - Maybe leaving blankc is fine idk

#### Next steps after making changes

Prompt copilot to take the whole sheet and write a code that will create the exact same output as the current parser worked on but will input every specimen from the collection descriptions.






