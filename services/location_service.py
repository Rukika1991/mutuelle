import pandas as pd
import re
import unicodedata

from config.settings import (
    LOCATION_MAPPING_FILE,
    EXPECTED_SECTORS,
    EXPECTED_CELLS,
    EXPECTED_VILLAGES,
)


# ============================================================
# BASIC TEXT CLEANING
# ============================================================

def clean_text(value):
    """
    Clean normal display text.
    Keeps the original readable form.
    """

    if pd.isna(value):
        return ""

    value = str(value)

    # Replace non-breaking spaces
    value = value.replace("\u00a0", " ")

    # Remove leading/trailing spaces
    value = value.strip()

    # Replace multiple spaces with one space
    value = re.sub(r"\s+", " ", value)

    return value


# ============================================================
# STRONG MATCHING KEY
# ============================================================

def match_key(value):
    """
    Create a strong key for location matching.

    Ignores:
    - Capital/lowercase
    - Extra spaces
    - Hyphens
    - Underscores
    - Dots
    - Apostrophes
    - Accents

    Example:

    KANYIRABIGOG
    Kanyirabigogo
    kanyira-bigogo

    -> KANYIRABIGOGO
    """

    value = clean_text(value)

    # Remove accents
    value = unicodedata.normalize("NFKD", value)

    value = "".join(
        character
        for character in value
        if not unicodedata.combining(character)
    )

    # Convert to uppercase
    value = value.upper()

    # Remove spaces and common separators
    value = re.sub(
        r"[\s\-_.'’/]+",
        "",
        value
    )

    # Keep only A-Z and numbers
    value = re.sub(
        r"[^A-Z0-9]",
        "",
        value
    )

    return value


# ============================================================
# NORMALIZE COLUMN NAMES
# ============================================================

def clean_column_name(value):

    value = str(value).strip().lower()

    value = re.sub(
        r"[\s\-/]+",
        "_",
        value
    )

    value = re.sub(
        r"[^a-z0-9_]",
        "",
        value
    )

    return value


def normalize_columns(df):

    df = df.copy()

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    return df


# ============================================================
# MANUAL LOCATION CORRECTIONS
# ============================================================

MANUAL_LOCATION_MAPPING = {

    "sector": {},

    "cell": {

        # Kanyirabigog and Kanyirabigogo
        "KANYIRABIGOG": "Kanyirabigogo",
        "KANYIRABIGOGO": "Kanyirabigogo",

        # Jenda and Jende
        "JENDA": "Jende",
        "JENDE": "Jende",
    },

    "village": {}
}


# ============================================================
# LOAD LOCATION MAPPING EXCEL
# ============================================================

def load_location_mapping():

    if not LOCATION_MAPPING_FILE.exists():

        return pd.DataFrame()

    try:

        mapping = pd.read_excel(
            LOCATION_MAPPING_FILE
        )

        mapping = normalize_columns(mapping)

        return mapping

    except Exception as e:

        print(
            f"WARNING: Could not load location mapping: {e}"
        )

        return pd.DataFrame()


# ============================================================
# BUILD EXPLICIT MAPPING
# ============================================================

def build_explicit_mapping(mapping):

    mappings = {

        "sector": {},
        "cell": {},
        "village": {}
    }

    # --------------------------------------------------------
    # ADD MANUAL MAPPING FIRST
    # --------------------------------------------------------

    for level, values in MANUAL_LOCATION_MAPPING.items():

        for original, canonical in values.items():

            mappings[level][
                match_key(original)
            ] = canonical

    # --------------------------------------------------------
    # IF EXCEL MAPPING IS EMPTY
    # --------------------------------------------------------

    if mapping.empty:

        return mappings

    # --------------------------------------------------------
    # FIND COLUMN NAMES
    # --------------------------------------------------------

    level_col = None
    original_col = None
    canonical_col = None

    for col in mapping.columns:

        key = match_key(col)

        if key in [
            "LEVEL",
            "LOCATIONLEVEL"
        ]:

            level_col = col

        elif key in [
            "ORIGINALNAME",
            "ORIGINAL",
            "SOURCENAME"
        ]:

            original_col = col

        elif key in [
            "CANONICALNAME",
            "CANONICAL",
            "STANDARDNAME"
        ]:

            canonical_col = col

    # --------------------------------------------------------
    # VALIDATE REQUIRED COLUMNS
    # --------------------------------------------------------

    if (
        level_col is None
        or original_col is None
        or canonical_col is None
    ):

        return mappings

    # --------------------------------------------------------
    # BUILD MAPPING
    # --------------------------------------------------------

    for _, row in mapping.iterrows():

        level = clean_text(
            row[level_col]
        ).lower()

        original = clean_text(
            row[original_col]
        )

        canonical = clean_text(
            row[canonical_col]
        )

        if (
            level in mappings
            and original
            and canonical
        ):

            mappings[level][
                match_key(original)
            ] = canonical

    return mappings


# ============================================================
# CREATE OFFICIAL MASTER LOOKUP
# ============================================================

def create_master_lookup(master):

    master = master.copy()

    lookup = {

        "sector": {},
        "cell": {},
        "village": {}
    }

    # --------------------------------------------------------
    # CLEAN LOCATION COLUMNS
    # --------------------------------------------------------

    required_columns = [
        "sector",
        "cell",
        "village"
    ]

    for column in required_columns:

        if column not in master.columns:

            raise ValueError(
                f"Master baseline missing column: {column}"
            )

        master[column] = (
            master[column]
            .apply(clean_text)
        )

    # --------------------------------------------------------
    # SECTORS
    # --------------------------------------------------------

    sectors = (
        master["sector"]
        .dropna()
        .apply(clean_text)
        .unique()
    )

    for sector in sectors:

        if sector:

            lookup["sector"][
                match_key(sector)
            ] = sector

    # --------------------------------------------------------
    # CELLS
    # --------------------------------------------------------

    cells = (
        master[
            ["sector", "cell"]
        ]
        .dropna()
        .drop_duplicates()
    )

    for _, row in cells.iterrows():

        sector = clean_text(
            row["sector"]
        )

        cell = clean_text(
            row["cell"]
        )

        key = (
            match_key(sector),
            match_key(cell)
        )

        lookup["cell"][key] = cell

    # --------------------------------------------------------
    # VILLAGES
    # --------------------------------------------------------

    villages = (
        master[
            ["sector", "cell", "village"]
        ]
        .dropna()
        .drop_duplicates()
    )

    for _, row in villages.iterrows():

        sector = clean_text(
            row["sector"]
        )

        cell = clean_text(
            row["cell"]
        )

        village = clean_text(
            row["village"]
        )

        key = (
            match_key(sector),
            match_key(cell),
            match_key(village)
        )

        lookup["village"][key] = village

    return lookup


# ============================================================
# APPLY LOCATION NORMALIZATION
# ============================================================

def normalize_locations(
    df,
    master_lookup,
    mapping=None
):

    df = df.copy()

    # --------------------------------------------------------
    # CLEAN LOCATION COLUMNS
    # --------------------------------------------------------

    location_columns = [
        "district",
        "sector",
        "cell",
        "village"
    ]

    for column in location_columns:

        if column in df.columns:

            df[column] = (
                df[column]
                .apply(clean_text)
            )

    # --------------------------------------------------------
    # BUILD EXPLICIT MAPPING
    # --------------------------------------------------------

    if mapping is None:

        mapping = pd.DataFrame()

    explicit_mapping = build_explicit_mapping(
        mapping
    )

    # ========================================================
    # STEP 1: NORMALIZE SECTOR
    # ========================================================

    if "sector" in df.columns:

        def normalize_sector(value):

            value = clean_text(value)

            key = match_key(value)

            # Apply explicit mapping
            if key in explicit_mapping["sector"]:

                value = explicit_mapping["sector"][key]

                key = match_key(value)

            # Match official master
            return master_lookup["sector"].get(
                key,
                value
            )

        df["sector"] = (
            df["sector"]
            .apply(normalize_sector)
        )

    # ========================================================
    # STEP 2: NORMALIZE CELL
    # ========================================================

    if (
        "sector" in df.columns
        and "cell" in df.columns
    ):

        def normalize_cell(row):

            sector = clean_text(
                row["sector"]
            )

            cell = clean_text(
                row["cell"]
            )

            cell_key = match_key(cell)

            # Apply explicit mapping
            if cell_key in explicit_mapping["cell"]:

                cell = explicit_mapping["cell"][cell_key]

            key = (
                match_key(sector),
                match_key(cell)
            )

            # Match official master
            return master_lookup["cell"].get(
                key,
                cell
            )

        df["cell"] = df.apply(
            normalize_cell,
            axis=1
        )

    # ========================================================
    # STEP 3: NORMALIZE VILLAGE
    # ========================================================

    if (
        "sector" in df.columns
        and "cell" in df.columns
        and "village" in df.columns
    ):

        def normalize_village(row):

            sector = clean_text(
                row["sector"]
            )

            cell = clean_text(
                row["cell"]
            )

            village = clean_text(
                row["village"]
            )

            village_key = match_key(village)

            # Apply explicit mapping
            if village_key in explicit_mapping["village"]:

                village = explicit_mapping[
                    "village"
                ][village_key]

            key = (
                match_key(sector),
                match_key(cell),
                match_key(village)
            )

            # Match official master
            return master_lookup["village"].get(
                key,
                village
            )

        df["village"] = df.apply(
            normalize_village,
            axis=1
        )

    return df


# ============================================================
# VALIDATE OFFICIAL RUBAVU STRUCTURE
# ============================================================

def validate_rubavu_structure(df):

    required_columns = [
        "sector",
        "cell",
        "village"
    ]

    missing = [

        column

        for column in required_columns

        if column not in df.columns
    ]

    if missing:

        return {

            "valid": False,

            "message": (
                "Missing location columns: "
                + ", ".join(missing)
            )
        }

    # --------------------------------------------------------
    # COUNT SECTORS
    # --------------------------------------------------------

    sectors = (
        df["sector"]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    # --------------------------------------------------------
    # COUNT UNIQUE CELLS
    # --------------------------------------------------------

    cells = (

        df[
            ["sector", "cell"]
        ]

        .replace("", pd.NA)

        .dropna()

        .drop_duplicates()

        .shape[0]
    )

    # --------------------------------------------------------
    # COUNT UNIQUE VILLAGES
    # --------------------------------------------------------

    villages = (

        df[
            ["sector", "cell", "village"]
        ]

        .replace("", pd.NA)

        .dropna()

        .drop_duplicates()

        .shape[0]
    )

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    valid = (

        sectors == EXPECTED_SECTORS
        and cells == EXPECTED_CELLS
        and villages == EXPECTED_VILLAGES
    )

    return {

        "valid": valid,

        "sectors": sectors,

        "cells": cells,

        "villages": villages,

        "expected_sectors": EXPECTED_SECTORS,

        "expected_cells": EXPECTED_CELLS,

        "expected_villages": EXPECTED_VILLAGES,
    }


# ============================================================
# SHOW STRUCTURE VALIDATION
# ============================================================

def structure_summary(validation):

    return (
        f"Sectors: {validation.get('sectors', 0)} "
        f"/ {EXPECTED_SECTORS}\n"

        f"Cells: {validation.get('cells', 0)} "
        f"/ {EXPECTED_CELLS}\n"

        f"Villages: {validation.get('villages', 0)} "
        f"/ {EXPECTED_VILLAGES}"
    )