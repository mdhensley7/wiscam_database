
import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine, text
from sqlalchemy.types import Text, Float, Date

# ✅ Connection string with your credentials
engine = create_engine(
    "postgresql+psycopg2://mdhensley7:KaylaHazel108510!@localhost:5432/great_wiscam_db",
    future=True
)

# ✅ Helper function to load CSV into Postgres
def load_csv(table_name, csv_path, dtype_map):
    df = pd.read_csv(csv_path)
    if "specimen_id" in df.columns:
        df["specimen_id"] = df["specimen_id"].astype(str)
    df.to_sql(table_name, engine, if_exists="append", index=False, dtype=dtype_map)

# ✅ Create parent table: specimens
with engine.begin() as conn:
    conn.execute(text("""
    CREATE TABLE IF NOT EXISTS specimens (
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
    """))

# ✅ Load specimens data
load_csv(
    "specimens",
    "specimens_clean.csv",
    {
        "specimen_id": Text(),
        "species": Text(), "sectionid": Text(), "collector": Text(),
        "hosttree1": Text(), "hosttree2": Text(), "county": Text(), "site": Text(),
        "lat": Float(), "long": Float(), "datecollected": Date()
    }
)

# ✅ Define child tables and their schema
tables = [
    ("pileus", "pileus_clean.csv", {
        "specimen_id": Text(), "species": Text(),
        "diameter_mm": Float(), "centercolor": Text(), "margincolor": Text(),
        "shape": Text(), "surfacetexture": Text(), "umbonate": Text(),
        "striationlength_mm": Float(), "appendiculate": Text(),
        "universalveilpresent": Text(), "staining": Text(), "contextcolor": Text()
    }),
    ("universal_veil_pileus", "universal_veil_pileus_clean.csv", {
        "specimen_id": Text(), "species": Text(),
        "form": Text(), "color": Text(), "texture": Text(),
        "attachment": Text(), "distribution": Text()
    }),
    ("lamellae", "lamellae_clean.csv", {
        "specimen_id": Text(), "species": Text(),
        "attachment": Text(), "colorinmasse": Text(), "distribution": Text(),
        "staining": Text(), "breadth": Text(), "edge": Text(), "othercomment": Text()
    }),
    ("lamellulae", "lamellulae_clean.csv", {
        "specimen_id": Text(), "species": Text(), "form": Text(),
        "amount": Text(), "distribution": Text(), "end": Text()
    }),
    ("stipe", "stipe_clean.csv", {
        "specimen_id": Text(), "species": Text(),
        "length_mm": Float(), "width_mm": Float(), "color": Text(), "shape": Text(),
        "decoration_bottom": Text(), "decoration_top": Text(), "staining": Text(),
        "annulus": Text(), "universal_veil": Text(), "context_type": Text(),
        "context_color": Text(), "context_stain": Text()
    }),
    ("annulus", "annulus_clean.csv", {
        "specimen_id": Text(), "species": Text(),
        "position": Text(), "form": Text(), "color": Text(), "staining": Text(),
        "remainingpercent": Float(), "condition": Text()
    }),
    ("universal_veil_stipe_base", "universal_veil_stipe_base_clean.csv", {
        "specimen_id": Text(), "species": Text(), "type": Text(),
        "texture": Text(), "color": Text(), "layered": Text(), "toughorflimsy": Text()
    }),
    ("basal_bulb", "basal_bulb_clean.csv", {
        "specimen_id": Text(), "species": Text(),
        "length_mm": Float(), "width_mm": Float(), "shape": Text()
    }),
]

# ✅ Create child tables and load data
with engine.begin() as conn:
    for table_name, csv_path, dtype_map in tables:
        # Create table with all columns dynamically
        columns_sql = ", ".join([f"{col} TEXT" for col in dtype_map.keys() if col != "specimen_id"])
        conn.execute(text(f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
          specimen_id TEXT REFERENCES specimens(specimen_id) ON DELETE CASCADE,
          {columns_sql}
        );
        """))
        # Load data
        load_csv(table_name, csv_path, dtype_map)
        # Add index for faster joins
        conn.execute(text(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_specimen_id ON {table_name}(specimen_id);"))
