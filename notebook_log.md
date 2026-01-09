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
python parse_specimen_form.py "/Users/hensley/Desktop/wiscam_database/normalizing_data/test_data/30181_test2.xlsx" "/Users/hensley/Desktop/wiscam_database/normalizing_data/test_data/30181_test2_normalized.xlsx" #the full path is overboard since all should be in the same folder but it should work either way
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

    - Maybe leaving blank is fine idk
    - Actually I think it is fine. It will be easy to fill those in manually

#### Next steps after making changes

Prompt copilot to take the whole sheet and write a code that will create the exact same output as the current parser worked on but will input every specimen from the collection descriptions.

## Converting data set to level 3 normalization

### Dec. 30 2025
Because a lot of the above steps really have to be done manually, I just prompted copilot to produce a script that uses the current parser we have but iterates it across all sheets in a single work book. The code for that can be found "batch_parse_specimen_workbook.py"

Let's see how well it works. Command below to initiate it:

```zsh

# Batch run across 100+ sheets:
python batch_parse_specimen_workbook.py collec_desc.xlsx -m parse_specimen_form_batch_ready --func parse_sheet -o specimens


```

From this, I got a normalized dataset called 'specimens_long' but it has some manual work that needs to be done to it to make sure it is a clean and complete dataset.

### Jan. 7 2026
Happy New Year!

In the interim, I had the code 'parse_specimen_form_batch_ready.py' normalize my dataset but my initial dataset was in pretty rough shape so I have spent the last day and a half cleaning the new dataset and making sure it is consistent and workable. 

I then prompted copilot to cut out any specimens that don't yet have a species id because I am going to use the ones that already do to start troubleshooting a relational database with the intention of being able to add specimens back in once I have a better idea of what their species id is.
#### Here is the code to pull specimens with species ids out into a new workbook

```zsh
python filter_species_out.py
```
The dataset I will be working on is 'samp_norm_data.xlsx' which I will copy into a new folder 'database'

```zsh
cd /Users/hensley/Desktop/wiscam_database
mkdir database
cp /Users/hensley/Desktop/wiscam_database/normalizing_data/samp_norm_data.xlsx /Users/hensley/Desktop/wiscam_database/database
```

# Building a relational database

I am going to use 'PostgreSQL' as the software to build the database because it has scalability and can be migrated to other hosting services if needed. I may just host it on my own server because I don't think my database is that intensive but might be nice to have it offered up to whoever else may want to do this work in the future.

## Downloading and installing PostgreSQL

I believe I have homebrew so I am going to try to intall it that way

```zsh
brew install postgresql@15
brew services start postgresql@15
brew info postgresql@15 #checks where home brew installed it
```

When installing via something like homebrew the binaries (commands??) are placed in a directory that might not be your actual path so you have to fix that by giving your shell (zsh) access to the binaries

```zsh
nano ~/.zshrc
#Once inside the shell file add this line
export PATH="/opt/homebrew/opt/postgresql@15/bin:$PATH"
#Save and exit the file and run below in the terminal to check to see if it work
source ~/.zshrc
```

Well that worked! Next is accessing Postgresql

```zsh
psql postgres
```

once in postgres I entered this code:
postgres=# CREATE USER mdhensley7 WITH PASSWORD 'KaylaHazel108510!';
postgres=# ALTER USER mdhensley7 CREATEDB;
postgres=# CREATE DATABASE great_wiscam_db OWNER mdhensley7;
postgres-# \q

';' is required at the end of code to execute SQL commands

Then I access the database via
```zsh
psql -U mdhensley7 -d great_wiscam_dbs
```
exit databade using \q

I am now going to link the database to python to make it a little easier to work with

```zsh
cd ~/wiscam_database/database

python -m venv .venv #Creates virtual environment in folder which reduce across computer conflict
source .venv/bin/activate
pip install pandas sqlalchemy psycopg2 openpyxl #install dependencies

```

### Jan 8 2026
I have been having crossover errors where conda is autoactivating my (WiscAmPhen) environment so I had to delete the conda initialize code in ~/.zshrc (the shell). IMPORTANT NOTE:
- if you need to activate a conda environment use this code:

```zsh

# Load Conda into your shell
source /path/to/miniconda3/etc/profile.d/conda.sh

# Then activate an environment
conda activate myenv #myenv I believe is replaced with environment name

```

Anywho, now that it is deactivated I am going to try and see if the code copilot gave me can in fact transfer the normalized data in the Excel sheet to postgresql in an organized manner. It seems like it lines up quite nicely so we will see if it can do it seemlessly.

I have been trying to get python and postgres to integrate but it has been wierd and I think I'll be better off manually entering the tables in pgadmin so I have a good understanding of the structure I am giving the database. 

That said pgadmin kinda sucks. It's slow and I have to run the whole query rather than in chunks which means I have to erase the code I used to generate tables etc. So I will be placing the code I used to create the database here even though it was being executed through pgadmin's query window.

#### Code used in pgadmin to set up tables for my database

```SQL
-- This is the parent table which the other tables with the morphological features will reference back to
CREATE TABLE specimens (
    specimen_id   TEXT PRIMARY KEY,
    species       TEXT,
    sectionid     TEXT,
    collector     TEXT,
    hosttree1     TEXT,
    hosttree2     TEXT,
    lat           DOUBLE PRECISION,
    long          DOUBLE PRECISION,
    county        TEXT,
    site          TEXT,
    datecollected DATE
);

-- Use code below to check if table has been made and is formatted correctly
SELECT column_name, data_type, is_nullable
FROM information_schema.columns
WHERE table_schema = 'public' AND table_name = 'specimens'
ORDER BY ordinal_position;
```

The next code block will include all of the child tables

```SQL
---
-- Child table 1: Pileus
---
CREATE TABLE pileus (
    specimen_id          TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species              TEXT,
    diameter_mm          DOUBLE PRECISION,
    centercolor          TEXT,
    margincolor          TEXT,
    shape                TEXT,
    surfacetexture       TEXT,
    umbonate             TEXT,
    striationlength_mm   DOUBLE PRECISION,
    appendiculate        TEXT,
    universalveilpresent TEXT,
    staining             TEXT,
    contextcolor         TEXT
);
CREATE INDEX IF NOT EXISTS idx_pileus_specimen_id ON pileus(specimen_id);

---
-- Child table 2: Universal_veil_pileus
---
CREATE TABLE universal_veil_pileus (
    specimen_id TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species     TEXT,
    form        TEXT,
    color       TEXT,
    texture     TEXT,
    attachment  TEXT,
    distribution TEXT
);
CREATE INDEX IF NOT EXISTS idx_universal_veil_pileus_specimen_id ON universal_veil_pileus(specimen_id);

---
-- Child table 3: Lamellae
---
CREATE TABLE lamellae (
    specimen_id   TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species       TEXT,
    attachment    TEXT,
    colorinmasse  TEXT,
    distribution  TEXT,
    staining      TEXT,
    breadth       TEXT,
    edge          TEXT,
    othercomment  TEXT
);
CREATE INDEX IF NOT EXISTS idx_lamellae_specimen_id ON lamellae(specimen_id);

---
-- Child table 4: Lamellulae
---
CREATE TABLE lamellulae (
    specimen_id  TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species      TEXT,
    form         TEXT,
    amount       TEXT,
    distribution TEXT,
    ending       TEXT
);
CREATE INDEX IF NOT EXISTS idx_lamellulae_specimen_id ON lamellulae(specimen_id);

---
-- Child table 5: Stipe
---
CREATE TABLE stipe (
    specimen_id      TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species          TEXT,
    length_mm        DOUBLE PRECISION,
    width_mm         DOUBLE PRECISION,
    color            TEXT,
    shape            TEXT,
    decoration_bottom TEXT,
    decoration_top    TEXT,
    staining         TEXT,
    annulus          TEXT,
    universal_veil   TEXT,
    context_type     TEXT,
    context_color    TEXT,
    context_stain    TEXT
);
CREATE INDEX IF NOT EXISTS idx_stipe_specimen_id ON stipe(specimen_id);

---
-- Child table 6: Annulus
---
CREATE TABLE annulus (
    specimen_id       TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species           TEXT,
    position          TEXT,
    form              TEXT,
    color             TEXT,
    staining          TEXT,
    remainingpercent  DOUBLE PRECISION,
    condition         TEXT
);
CREATE INDEX IF NOT EXISTS idx_annulus_specimen_id ON annulus(specimen_id);

---
-- Child table 7: Universal_veil_stipe_base
---
CREATE TABLE universal_veil_stipe_base (
    specimen_id   TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species       TEXT,
    type          TEXT,
    texture       TEXT,
    color         TEXT,
    layered       TEXT,
    toughorflimsy TEXT
);
CREATE INDEX IF NOT EXISTS idx_universal_veil_stipe_base_specimen_id ON universal_veil_stipe_base(specimen_id);

---
-- Child table 8: basal_bulb
---
CREATE TABLE basal_bulb (
    specimen_id TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
    species     TEXT,
    length_mm   DOUBLE PRECISION,
    width_mm    DOUBLE PRECISION,
    shape       TEXT
);
CREATE INDEX IF NOT EXISTS idx_basal_bulb_specimen_id ON basal_bulb(specimen_id);

```

All the tables are into the data base and I was even able to configure the 'specimens' csv sheet so that it would import into the table. What's cool is that it totally worked and I can now start to see how this all will come together.

### Next steps

- Figure out how to make importing and editing data easy
    - This likely means having all of your columns completely set and defined in each table
    - We may have to add a few more columns but maybe there's and easy way to add columns and only have to import that data.

- Part of this is keeping in mind this is not the full dataset. If you are making edits, do it to the complete excel sheet "specimens_long.xlsx" so that formatting remains the same across the board.
    - I just editted the hosttree names in only the specimen_long_filtered.xlsx which means I will have to go back to the main excel workbook to apply changes there as well.

- I think it would be a good call to feed copilot the whole sheet and tell it to fix any formatting inconsistencies on the whole sheet and clean anything else so then we can import .csv files that are already clean. On top of that, ask copilot to add an excel spreadsheet to the workbook that explains how to input data so that it remains clean as it gets fed into the database.