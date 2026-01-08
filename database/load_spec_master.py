
# load_specimens_species_master.py
import os
import sys
import pandas as pd
from datetime import datetime
from sqlalchemy import (
    create_engine, MetaData, Table, Column, Integer, BigInteger, String, Text,
    Date, ForeignKey, DECIMAL
)
from sqlalchemy.dialects.postgresql import insert

# ----------------------------
# CONFIG
# ----------------------------
# Use env vars or fallback to local defaults
PG_USER = os.getenv("PG_USER", "mdhensley7")          # your Postgres role
PG_PASS = os.getenv("PG_PASS", "KaylaHazel108510!")                    # set if needed; empty for local trust auth
PG_HOST = os.getenv("PG_HOST", "localhost")
PG_PORT = os.getenv("PG_PORT", "5432")
PG_DB   = os.getenv("PG_DB",   "great_wiscam_db")

# If your local auth does not use a password, use this URI without PG_PASS:
if PG_PASS:
    DATABASE_URL = f"postgresql+psycopg2://{PG_USER}:{PG_PASS}@{PG_HOST}:{PG_PORT}/{PG_DB}"
else:
    DATABASE_URL = f"postgresql+psycopg2://{PG_USER}@{PG_HOST}:{PG_PORT}/{PG_DB}"

# Path to the Excel workbook (change if needed)
EXCEL_PATH = "specimens_long_filtered.xlsx"

# Treat these as "missing"
EMPTY_STRINGS = {"", "na", "n/a", "null", "none", "NA", "N/A", "None"}

# ----------------------------
# DB METADATA & TABLES
# ----------------------------
engine = create_engine(DATABASE_URL, future=True)
metadata = MetaData()

# Species master (surrogate PK; unique species_name for UPSERT)
species_tbl = Table(
    "species", metadata,
    Column("species_id", Integer, primary_key=True, autoincrement=True),
    Column("species_name", Text, nullable=False, unique=True),
    Column("section_name", Text),   # optional convenience field
    Column("notes", Text)
)

# Specimens (FK -> species)
specimens_tbl = Table(
    "specimens", metadata,
    Column("specimenid", String, primary_key=True),  # matches workbook IDs (keep as text)
    Column("species_id", Integer, ForeignKey("species.species_id", onupdate="CASCADE"), nullable=False),
    Column("sectionid", String),
    Column("collector", String),
    Column("hosttree1", String),
    Column("hosttree2", String),
    Column("lat", DECIMAL(10, 6)),
    Column("long", DECIMAL(10, 6)),
    Column("county", String),
    Column("site", String),
    Column("datecollected", Date)
)

# Helper: child table with surrogate PK & FK to specimens
def child_table(name, extra_cols):
    cols = [
        Column(f"{name}_id", BigInteger, primary_key=True, autoincrement=True),
        Column("specimenid", String, ForeignKey("specimens.specimenid", ondelete="CASCADE"), nullable=False),
    ]
    cols.extend(extra_cols)
    return Table(name, metadata, *cols)

# Child tables (keep 'species' text column for auditing; joins should use specimens/species_id)
pileus_tbl = child_table("pileus", [
    Column("species", Text),
    Column("diameter_mm", Integer),
    Column("centercolor", Text),
    Column("margincolor", Text),
    Column("shape", Text),
    Column("surfacetexture", Text),
    Column("umbonate", Text),
    Column("striationlength_mm", Integer),
    Column("appendiculate", Text),
    Column("universalveilpresent", Text),
    Column("staining", Text),
    Column("contextcolor", Text),
])

univ_veil_pileus_tbl = child_table("universal_veil_pileus", [
    Column("species", Text),
    Column("form", Text),
    Column("color", Text),
    Column("texture", Text),
    Column("attachment", Text),
    Column("distribution", Text),
])

lamellae_tbl = child_table("lamellae", [
    Column("species", Text),
    Column("attachment", Text),
    Column("colorinmasse", Text),
    Column("distribution", Text),
    Column("staining", Text),
    Column("breadth", Text),
    Column("edge", Text),
    Column("othercomment", Text),
])

lamellulae_tbl = child_table("lamellulae", [
    Column("species", Text),
    Column("form", Text),
    Column("amount", Text),
    Column("distribution", Text),
    Column("end", Text),
])

stipe_tbl = child_table("stipe", [
    Column("species", Text),
    Column("length_mm", Integer),
    Column("width_mm", Integer),
    Column("color", Text),
    Column("shape", Text),
    Column("decoration_bottom", Text),
    Column("decoration_top", Text),
    Column("staining", Text),
    Column("annulus", Text),
    Column("universal_veil", Text),
    Column("context_type", Text),
    Column("context_color", Text),
    Column("context_stain", Text),
])

annulus_tbl = child_table("annulus", [
    Column("species", Text),
    Column("position", Text),
    Column("form", Text),
    Column("color", Text),
    Column("staining", Text),
    Column("remainingpercent", Text),
    Column("condition", Text),
])

univ_veil_stipe_base_tbl = child_table("universal_veil_stipe_base", [
    Column("species", Text),
    Column("type", Text),
    Column("texture", Text),
    Column("color", Text),
    Column("layered", Text),
    Column("toughorflimsy", Text),
])

basal_bulb_tbl = child_table("basal_bulb", [
    Column("species", Text),
    Column("length_mm", Integer),
    Column("width_mm", Integer),
    Column("shape", Text),
])

# Create all tables if they don't exist
metadata.create_all(engine)

# ----------------------------
# HELPERS
# ----------------------------
def normalize_value(x):
    if x is None:
        return None
    s = str(x).strip()
    return None if s.lower() in EMPTY_STRINGS else s

def coerce_int(x):
    s = normalize_value(x)
    if s is None:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None

def coerce_decimal(x):
    s = normalize_value(x)
    if s is None:
        return None
    try:
        return float(s)
    except ValueError:
        return None

def parse_date(x):
    s = normalize_value(x)
    if s is None:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(s).date()
    except Exception:
        return None

def upsert_pk(table, rows, pk_col):
    """UPSERT using the primary key column."""
    if not rows:
        return
    with engine.begin() as conn:
        stmt = insert(table).values(rows)
        update_cols = {c.name: stmt.excluded.get(c.name) for c in table.columns if c.name != pk_col}
        stmt = stmt.on_conflict_do_update(index_elements=[pk_col], set_=update_cols)
        conn.execute(stmt)

def upsert_unique(table, rows, unique_cols):
    """UPSERT using a unique constraint (e.g., species_name)."""
    if not rows:
        return
    with engine.begin() as conn:
        stmt = insert(table).values(rows)
        update_cols = {c.name: stmt.excluded.get(c.name) for c in table.columns if c.name not in unique_cols}
        stmt = stmt.on_conflict_do_update(index_elements=unique_cols, set_=update_cols)
        conn.execute(stmt)

def delete_existing_specimen_rows(table, specimen_ids):
    """Make sheet loads idempotent by deleting existing rows for the specimen set, then re-insert."""
    if not specimen_ids:
        return
    with engine.begin() as conn:
        # Use parameterized tuple for IN clause
        conn.execute(table.delete().where(table.c.specimenid.in_(specimen_ids)))

def get_specimenid(series_or_row_dict):
    """Robustly fetch specimenid from either 'specimenid' or 'specimen_id' (text)."""
    if isinstance(series_or_row_dict, dict):
        val = series_or_row_dict.get("specimenid")
        if val is None:
            val = series_or_row_dict.get("specimen_id")
        return normalize_value(val)
    else:
        # pandas Series
        val = series_or_row_dict.get("specimenid")
        if pd.isna(val):
            val = series_or_row_dict.get("specimen_id")
        return normalize_value(val)

# ----------------------------
# LOAD EXCEL
# ----------------------------
if not os.path.exists(EXCEL_PATH):
    print(f"ERROR: Excel file not found: {EXCEL_PATH}", file=sys.stderr)
    sys.exit(1)

xl = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")

# -------- 1) SPECIES MASTER --------
specimens_df = xl.parse("specimens", dtype=str).rename(columns=str.lower)
specimens_df = specimens_df.dropna(how="all")  # drop fully empty rows

# Distinct species names + optional section names from the specimens sheet
species_rows = []
for _, r in specimens_df.iterrows():
    sp = normalize_value(r.get("species"))
    if not sp:
        continue
    section = normalize_value(r.get("sectionid"))
    species_rows.append({"species_name": sp, "section_name": section})

# Deduplicate species rows by species_name (keep last section_name seen)
if species_rows:
    df_sp = pd.DataFrame(species_rows).drop_duplicates(subset=["species_name"], keep="last")
    upsert_unique(species_tbl, df_sp.to_dict(orient="records"), unique_cols=["species_name"])

# Build map: species_name -> species_id
from sqlalchemy import select

with engine.begin() as conn:
    result = conn.execute(select(species_tbl.c.species_name, species_tbl.c.species_id))
    species_map = {r["species_name"]: r["species_id"] for r in result.mappings()}

# -------- 2) SPECIMENS --------
specimen_rows = []
for _, r in specimens_df.iterrows():
    specimenid = get_specimenid(r)
    species_name = normalize_value(r.get("species"))
    if not specimenid or not species_name:
        continue
    species_id = species_map.get(species_name)
    specimen_rows.append({
        "specimenid": specimenid,
        "species_id": species_id,
        "sectionid": normalize_value(r.get("sectionid")),
        "collector": normalize_value(r.get("collector")),
        "hosttree1": normalize_value(r.get("hosttree1")),
        "hosttree2": normalize_value(r.get("hosttree2")),
        "lat": coerce_decimal(r.get("lat")),
        "long": coerce_decimal(r.get("long")),
        "county": normalize_value(r.get("county")),
        "site": normalize_value(r.get("site")),
        "datecollected": parse_date(r.get("datecollected")),
    })

# UPSERT specimens by PK
upsert_pk(specimens_tbl, specimen_rows, pk_col="specimenid")

# -------- 3) CHILD TABLES --------
def load_child(sheet_name, table, int_cols=None, dec_cols=None):
    df = xl.parse(sheet_name, dtype=str).rename(columns=str.lower)
    df = df.dropna(how="all")  # drop fully empty rows
    int_cols = int_cols or []
    dec_cols = dec_cols or []

    # Prepare rows and collect specimenids to clear for idempotency
    rows = []
    specimen_ids_to_clear = set()

    for _, r in df.iterrows():
        specimenid = get_specimenid(r)
        if not specimenid:
            continue
        specimen_ids_to_clear.add(specimenid)

        row = {"specimenid": specimenid}
        # Copy matching columns if they exist in this sheet
        for col in table.columns.keys():
            if col in ("specimenid", f"{table.name}_id"):
                continue  # handled / auto
            if col in df.columns:
                row[col] = normalize_value(r.get(col))

        # numeric coercions
        for c in int_cols:
            if c in row:
                row[c] = coerce_int(row[c])
        for c in dec_cols:
            if c in row:
                row[c] = coerce_decimal(row[c])

        rows.append(row)

    # Delete existing rows for these specimens, then insert fresh
    delete_existing_specimen_rows(table, list(specimen_ids_to_clear))
    if rows:
        with engine.begin() as conn:
            conn.execute(table.insert(), rows)

# Load each child (numeric coercions where applicable)
load_child("pileus", pileus_tbl, int_cols=["diameter_mm", "striationlength_mm"])
load_child("universal_veil_pileus", univ_veil_pileus_tbl)
load_child("lamellae", lamellae_tbl)
load_child("lamellulae", lamellulae_tbl)
load_child("stipe", stipe_tbl, int_cols=["length_mm", "width_mm"])
load_child("annulus", annulus_tbl)
load_child("universal_veil_stipe_base", univ_veil_stipe_base_tbl)
load_child("basal_bulb", basal_bulb_tbl, int_cols=["length_mm", "width_mm"])

print("Load complete ✅")
