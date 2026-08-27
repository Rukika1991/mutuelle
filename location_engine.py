import pandas as pd
import re
import unicodedata


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(value):
    """Basic text cleaning."""
    if pd.isna(value):
        return ""

    value = str(value)
    value = value.replace("\u00a0", " ")
    value = value.strip()
    value = re.sub(r"\s+", " ", value)

    return value


# ============================================================
# STRONG MATCH KEY
# ============================================================

def match_key(value):
    """
    Strong comparison key.

    Ignores:
    - upper/lower case
    - spaces
    - hyphens
    - underscores
    - dots
    - apostrophes
    - accents
    """

    value = clean_text(value)

    value = unicodedata.normalize("NFKD", value)

    value = "".join(
        character
        for character in value
        if not unicodedata.combining(character)
    )

    value = value.upper()

    value = re.sub(
        r"[\s\-_.'’/]+",
        "",
        value
    )

    value = re.sub(
        r"[^A-Z0-9]",
        "",
        value
    )

    return value


# ============================================================
# COLUMN NORMALIZATION
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

    "sector": {
    },

    "cell": {

        "KANYIRABIGOG":
            "Kanyirabigogo",

        "KANYIRABIGOGO":
            "Kanyirabigogo",

        "JENDA":
            "Jende",

        "JENDE":
            "Jende",
    },

    "village": {
    }
}


# ============================================================
# CREATE MASTER LOOKUP
# ============================================================

def create_master_lookup(master):

    master = master.copy()

    lookup = {
        "sector": {},
        "cell": {},
        "village": {}
    }

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

        sector = clean_text(row["sector"])
        cell = clean_text(row["cell"])

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

        sector = clean_text(row["sector"])
        cell = clean_text(row["cell"])
        village = clean_text(row["village"])

        key = (
            match_key(sector),
            match_key(cell),
            match_key(village)
        )

        lookup["village"][key] = village

    return lookup


# ============================================================
# EXPLICIT MAPPING
# ============================================================

def build_explicit_mapping():

    mappings = {
        "sector": {},
        "cell": {},
        "village": {}
    }

    for level, values in MANUAL_LOCATION_MAPPING.items():

        for original, canonical in values.items():

            mappings[level][
                match_key(original)
            ] = canonical

    return mappings


# ============================================================
# NORMALIZE LOCATIONS
# ============================================================

def normalize_locations(df, master_lookup):

    df = df.copy()

    explicit_mapping = build_explicit_mapping()

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
    # NORMALIZE SECTOR
    # --------------------------------------------------------

    if "sector" in df.columns:

        def normalize_sector(value):

            value = clean_text(value)

            key = match_key(value)

            if key in explicit_mapping["sector"]:

                value = explicit_mapping["sector"][key]

                key = match_key(value)

            return master_lookup["sector"].get(
                key,
                value
            )

        df["sector"] = (
            df["sector"]
            .apply(normalize_sector)
        )

    # --------------------------------------------------------
    # NORMALIZE CELL
    # --------------------------------------------------------

    if (
        "sector" in df.columns
        and "cell" in df.columns
    ):

        def normalize_cell(row):

            sector = clean_text(row["sector"])
            cell = clean_text(row["cell"])

            cell_key = match_key(cell)

            if cell_key in explicit_mapping["cell"]:

                cell = explicit_mapping["cell"][cell_key]

            key = (
                match_key(sector),
                match_key(cell)
            )

            return master_lookup["cell"].get(
                key,
                cell
            )

        df["cell"] = (
            df.apply(
                normalize_cell,
                axis=1
            )
        )

    # --------------------------------------------------------
    # NORMALIZE VILLAGE
    # --------------------------------------------------------

    if (
        "sector" in df.columns
        and "cell" in df.columns
        and "village" in df.columns
    ):

        def normalize_village(row):

            sector = clean_text(row["sector"])
            cell = clean_text(row["cell"])
            village = clean_text(row["village"])

            village_key = match_key(village)

            if village_key in explicit_mapping["village"]:

                village = (
                    explicit_mapping["village"][village_key]
                )

            key = (
                match_key(sector),
                match_key(cell),
                match_key(village)
            )

            return master_lookup["village"].get(
                key,
                village
            )

        df["village"] = (
            df.apply(
                normalize_village,
                axis=1
            )
        )

    return df


# ============================================================
# VALIDATE RUBAVU STRUCTURE
# ============================================================

def validate_structure(df):

    result = {}

    sectors = (
        df["sector"]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    cells = (
        df[
            ["sector", "cell"]
        ]
        .replace("", pd.NA)
        .dropna()
        .drop_duplicates()
        .shape[0]
    )

    villages = (
        df[
            [
                "sector",
                "cell",
                "village"
            ]
        ]
        .replace("", pd.NA)
        .dropna()
        .drop_duplicates()
        .shape[0]
    )

    result["sectors"] = sectors
    result["cells"] = cells
    result["villages"] = villages

    result["valid"] = (
        sectors == 12
        and cells == 80
        and villages == 519
    )

    return result


# ============================================================
# LOCATION ENGINE CLASS
# ============================================================

class LocationEngine:

    def __init__(self, baseline):

        self.baseline = baseline.copy()

        self.master_lookup = create_master_lookup(
            self.baseline
        )

    # --------------------------------------------------------
    # NORMALIZE
    # --------------------------------------------------------

    def normalize(self, df):

        return normalize_locations(
            df,
            self.master_lookup
        )

    # --------------------------------------------------------
    # VALIDATE
    # --------------------------------------------------------

    def validate(self, df):

        return validate_structure(df)

    # --------------------------------------------------------
    # PROCESS
    # --------------------------------------------------------

    def process(self, df):

        normalized = self.normalize(df)

        validation = self.validate(normalized)

        return normalized, validation

    # --------------------------------------------------------
    # OFFICIAL SECTORS
    # --------------------------------------------------------

    def get_sectors(self):

        return sorted(
            set(
                self.master_lookup["sector"].values()
            )
        )

    # --------------------------------------------------------
    # OFFICIAL CELLS
    #
    # IMPORTANT:
    # Cell identity is:
    #     sector + cell
    #
    # Therefore we return the actual cell mapping keys.
    # --------------------------------------------------------

    def get_cells(self):

        return list(
            self.master_lookup["cell"].items()
        )

    # --------------------------------------------------------
    # OFFICIAL VILLAGES
    #
    # IMPORTANT:
    # Village identity is:
    #     sector + cell + village
    #
    # Therefore we return the actual village mapping keys.
    # --------------------------------------------------------

    def get_villages(self):

        return list(
            self.master_lookup["village"].items()
        )