import pandas as pd
import re
import unicodedata
from pathlib import Path
from datetime import datetime
from rapidfuzz import fuzz, process


# ============================================================
# CONFIGURATION
# ============================================================

BASELINE_FILE = Path("data/baseline/RubavunewAugust.xlsx")
PROGRESS_DIR = Path("data/progress")
PROCESSED_DIR = Path("data/processed")
OUTPUT_DIR = Path("output")

MATCH_REPORT = OUTPUT_DIR / "village_matching_report.xlsx"
MAPPING_FILE = OUTPUT_DIR / "village_name_mapping.xlsx"
ALIAS_FILE = OUTPUT_DIR / "village_aliases.xlsx"


# ============================================================
# MATCHING SETTINGS
# ============================================================

# 95+ = automatically accepted fuzzy match
AUTO_MATCH_SCORE = 95

# 85-94 = review
REVIEW_SCORE = 85

# Below 85 = unmatched
MIN_MATCH_SCORE = 85


# ============================================================
# CREATE DIRECTORIES
# ============================================================

PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(value):

    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    # Remove accents
    value = unicodedata.normalize("NFKD", value)

    value = "".join(
        c for c in value
        if not unicodedata.combining(c)
    )

    # Replace punctuation with spaces
    value = re.sub(r"[^A-Z0-9]+", " ", value)

    # Remove multiple spaces
    value = re.sub(r"\s+", " ", value)

    return value.strip()


# ============================================================
# NORMALIZE COLUMN NAME
# ============================================================

def normalize_column_name(value):

    value = str(value).strip().lower()

    value = re.sub(
        r"[^a-z0-9]+",
        "_",
        value
    )

    return value.strip("_")


# ============================================================
# FIND GEOGRAPHIC COLUMNS
# ============================================================

def find_geo_columns(df):

    normalized = {
        normalize_column_name(c): c
        for c in df.columns
    }

    aliases = {

        "district": [
            "district",
            "district_name"
        ],

        "sector": [
            "sector",
            "sector_name"
        ],

        "cell": [
            "cell",
            "cell_name"
        ],

        "village": [
            "village",
            "village_name",
            "village_name_"
        ]
    }

    result = {}

    for standard, possible in aliases.items():

        found = None

        for candidate in possible:

            if candidate in normalized:
                found = normalized[candidate]
                break

        if found is None:

            for norm_name, original in normalized.items():

                if standard in norm_name:
                    found = original
                    break

        if found is None:

            raise ValueError(
                f"Could not find {standard} column. "
                f"Available columns: {list(df.columns)}"
            )

        result[standard] = found

    return result


# ============================================================
# READ BASELINE
# ============================================================

def read_baseline():

    print("\n" + "=" * 80)
    print("READING BASELINE MASTER")
    print("=" * 80)

    print(f"File: {BASELINE_FILE}")

    df = pd.read_excel(BASELINE_FILE)

    print(f"Baseline shape: {df.shape}")

    print("\nBaseline columns:")

    for column in df.columns:
        print(f" - {column}")

    geo = find_geo_columns(df)

    print("\nGeographic columns detected:")

    for key, value in geo.items():
        print(f" {key}: {value}")

    result = pd.DataFrame()

    result["district"] = df[geo["district"]]
    result["sector"] = df[geo["sector"]]
    result["cell"] = df[geo["cell"]]
    result["village"] = df[geo["village"]]

    return result


# ============================================================
# READ PROGRESS FILE
# ============================================================

def read_progress_file(file_path):

    print("\n" + "=" * 80)
    print("READING PROGRESS FILE")
    print("=" * 80)

    print(f"File: {file_path}")

    raw = pd.read_excel(
        file_path,
        header=None
    )

    print(f"Raw shape: {raw.shape}")

    # Progress report has 3 header rows
    df = raw.iloc[3:].copy()

    # First 19 columns
    df = df.iloc[:, :19]

    df.columns = [
        "district",
        "sector",
        "cell",
        "village",

        "level1_households",
        "level1_members",
        "level1_amount",

        "level2_households",
        "level2_members",
        "level2_amount",

        "level3_households",
        "level3_members",
        "level3_amount",

        "level4_households",
        "level4_members",
        "level4_amount",

        "level5_households",
        "level5_members",
        "level5_amount"
    ]

    return df


# ============================================================
# CLEAN GEOGRAPHIC DATA
# ============================================================

def prepare_geo_dataframe(df):

    df = df.copy()

    for column in [
        "district",
        "sector",
        "cell",
        "village"
    ]:

        df[column] = (
            df[column]
            .astype(str)
            .str.strip()
        )

        df[column] = df[column].replace(
            [
                "nan",
                "None",
                ""
            ],
            pd.NA
        )

    df = df[
        df["district"].notna()
        &
        df["sector"].notna()
        &
        df["cell"].notna()
        &
        df["village"].notna()
    ].copy()

    # Normalized versions
    df["district_norm"] = df["district"].apply(
        normalize_text
    )

    df["sector_norm"] = df["sector"].apply(
        normalize_text
    )

    df["cell_norm"] = df["cell"].apply(
        normalize_text
    )

    df["village_norm"] = df["village"].apply(
        normalize_text
    )

    return df


# ============================================================
# CREATE MASTER VILLAGES
# ============================================================

def create_master_villages(baseline):

    print("\n" + "=" * 80)
    print("CREATING MASTER VILLAGE LIST")
    print("=" * 80)

    master = prepare_geo_dataframe(
        baseline
    )

    master = master.drop_duplicates(
        subset=[
            "district_norm",
            "sector_norm",
            "cell_norm",
            "village_norm"
        ]
    ).copy()

    master["master_id"] = range(
        1,
        len(master) + 1
    )

    unique_cells = master[
        [
            "district_norm",
            "sector_norm",
            "cell_norm"
        ]
    ].drop_duplicates().shape[0]

    print(
        f"Master villages: {len(master):,}"
    )

    print(
        f"Master cells: {unique_cells:,}"
    )

    return master


# ============================================================
# LOAD PERMANENT ALIASES
# ============================================================

def load_aliases():

    if not ALIAS_FILE.exists():

        print("\nNo permanent alias file found.")
        print(
            f"Creating it later at: {ALIAS_FILE}"
        )

        return pd.DataFrame(
            columns=[
                "source_village",
                "standard_village",
                "source_cell",
                "standard_cell",
                "source_sector",
                "standard_sector",
                "status",
                "match_type",
                "match_score",
                "approved_by",
                "approved_date"
            ]
        )

    print("\n" + "=" * 80)
    print("LOADING PERMANENT ALIASES")
    print("=" * 80)

    aliases = pd.read_excel(
        ALIAS_FILE
    )

    print(
        f"Alias records loaded: {len(aliases):,}"
    )

    if "status" not in aliases.columns:
        aliases["status"] = "APPROVED"

    aliases["status"] = (
        aliases["status"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    return aliases


# ============================================================
# SAVE ALIASES
# ============================================================

def save_aliases(aliases):

    if aliases.empty:

        aliases = pd.DataFrame(
            columns=[
                "source_village",
                "standard_village",
                "source_cell",
                "standard_cell",
                "source_sector",
                "standard_sector",
                "status",
                "match_type",
                "match_score",
                "approved_by",
                "approved_date"
            ]
        )

    aliases = aliases.copy()

    # Remove duplicates
    aliases = aliases.drop_duplicates(
        subset=[
            "source_village",
            "standard_village",
            "source_cell",
            "standard_cell",
            "source_sector",
            "standard_sector"
        ],
        keep="last"
    )

    aliases = aliases.sort_values(
        [
            "source_sector",
            "source_cell",
            "source_village"
        ],
        na_position="last"
    )

    aliases.to_excel(
        ALIAS_FILE,
        index=False
    )

    print(
        f"Permanent alias file saved: {ALIAS_FILE}"
    )

    print(
        f"Total alias records: {len(aliases):,}"
    )


# ============================================================
# BUILD APPROVED ALIAS LOOKUPS
# ============================================================

def build_alias_lookup(aliases):

    village_aliases = {}
    cell_aliases = {}
    sector_aliases = {}

    if aliases.empty:
        return (
            village_aliases,
            cell_aliases,
            sector_aliases
        )

    approved = aliases[
        aliases["status"].astype(str).str.upper() == "APPROVED"
    ].copy()

    for _, row in approved.iterrows():

        # ----------------------------------------------------
        # Village alias
        # ----------------------------------------------------

        source_village = normalize_text(
            row.get("source_village", "")
        )

        standard_village = normalize_text(
            row.get("standard_village", "")
        )

        source_cell = normalize_text(
            row.get("source_cell", "")
        )

        standard_cell = normalize_text(
            row.get("standard_cell", "")
        )

        source_sector = normalize_text(
            row.get("source_sector", "")
        )

        standard_sector = normalize_text(
            row.get("standard_sector", "")
        )

        if source_village and standard_village:

            key = (
                source_sector,
                source_cell,
                source_village
            )

            village_aliases[key] = standard_village

        # ----------------------------------------------------
        # Cell alias
        # ----------------------------------------------------

        if source_cell and standard_cell:

            key = (
                source_sector,
                source_cell
            )

            cell_aliases[key] = standard_cell

        # ----------------------------------------------------
        # Sector alias
        # ----------------------------------------------------

        if source_sector and standard_sector:

            sector_aliases[
                source_sector
            ] = standard_sector

    print(
        f"Approved village aliases: "
        f"{len(village_aliases):,}"
    )

    print(
        f"Approved cell aliases: "
        f"{len(cell_aliases):,}"
    )

    print(
        f"Approved sector aliases: "
        f"{len(sector_aliases):,}"
    )

    return (
        village_aliases,
        cell_aliases,
        sector_aliases
    )


# ============================================================
# APPLY ALIASES
# ============================================================

def apply_aliases(
    row,
    village_aliases,
    cell_aliases,
    sector_aliases
):

    district = normalize_text(
        row["district"]
    )

    sector = normalize_text(
        row["sector"]
    )

    cell = normalize_text(
        row["cell"]
    )

    village = normalize_text(
        row["village"]
    )

    original_sector = sector
    original_cell = cell
    original_village = village

    # --------------------------------------------------------
    # Sector alias
    # --------------------------------------------------------

    if sector in sector_aliases:

        sector = sector_aliases[
            sector
        ]

    # --------------------------------------------------------
    # Cell alias
    # --------------------------------------------------------

    cell_key = (
        original_sector,
        original_cell
    )

    if cell_key in cell_aliases:

        cell = cell_aliases[
            cell_key
        ]

    else:

        # Try after sector normalization
        cell_key = (
            sector,
            original_cell
        )

        if cell_key in cell_aliases:

            cell = cell_aliases[
                cell_key
            ]

    # --------------------------------------------------------
    # Village alias
    # --------------------------------------------------------

    village_key = (
        original_sector,
        original_cell,
        original_village
    )

    if village_key in village_aliases:

        village = village_aliases[
            village_key
        ]

    else:

        village_key = (
            sector,
            cell,
            original_village
        )

        if village_key in village_aliases:

            village = village_aliases[
                village_key
            ]

    return {
        "district_norm": district,
        "sector_norm": sector,
        "cell_norm": cell,
        "village_norm": village,

        "alias_used": (
            original_sector != sector
            or
            original_cell != cell
            or
            original_village != village
        )
    }


# ============================================================
# CREATE MASTER LOOKUP
# ============================================================

def create_master_lookup(master):

    lookup = {}

    for _, row in master.iterrows():

        key = (
            row["district_norm"],
            row["sector_norm"],
            row["cell_norm"],
            row["village_norm"]
        )

        lookup[key] = row

    return lookup


# ============================================================
# BUILD CELL INDEX
# ============================================================

def build_cell_index(master):

    cell_index = {}

    for _, row in master.iterrows():

        key = (
            row["district_norm"],
            row["sector_norm"],
            row["cell_norm"]
        )

        if key not in cell_index:
            cell_index[key] = []

        cell_index[key].append(
            row["village_norm"]
        )

    return cell_index


# ============================================================
# FUZZY MATCH
# ============================================================

def fuzzy_match(
    candidates,
    village
):

    if not candidates:
        return None, 0

    result = process.extractOne(
        village,
        candidates,
        scorer=fuzz.ratio
    )

    if result is None:
        return None, 0

    matched_name = result[0]
    score = result[1]

    return matched_name, score


# ============================================================
# MATCH ONE PROGRESS FILE
# ============================================================

def match_progress_file(
    progress_file,
    master,
    cell_index,
    master_lookup,
    village_aliases,
    cell_aliases,
    sector_aliases
):

    print("\n" + "=" * 80)
    print(
        f"MATCHING: {progress_file.name}"
    )
    print("=" * 80)

    df = read_progress_file(
        progress_file
    )

    df = prepare_geo_dataframe(
        df
    )

    print(
        f"Progress villages: {len(df):,}"
    )

    # --------------------------------------------------------
    # Duplicate detection
    # --------------------------------------------------------

    duplicate_mask = df.duplicated(
        subset=[
            "district_norm",
            "sector_norm",
            "cell_norm",
            "village_norm"
        ],
        keep=False
    )

    df["duplicate_in_report"] = (
        duplicate_mask
    )

    results = []

    # ========================================================
    # MATCH EACH ROW
    # ========================================================

    for index, row in df.iterrows():

        original_district = row["district"]
        original_sector = row["sector"]
        original_cell = row["cell"]
        original_village = row["village"]

        # ----------------------------------------------------
        # Apply permanent aliases
        # ----------------------------------------------------

        normalized = apply_aliases(
            row,
            village_aliases,
            cell_aliases,
            sector_aliases
        )

        district = normalized[
            "district_norm"
        ]

        sector = normalized[
            "sector_norm"
        ]

        cell = normalized[
            "cell_norm"
        ]

        village = normalized[
            "village_norm"
        ]

        alias_used = normalized[
            "alias_used"
        ]

        # ----------------------------------------------------
        # EXACT MATCH
        # ----------------------------------------------------

        key = (
            district,
            sector,
            cell,
            village
        )

        exact = master_lookup.get(
            key
        )

        if exact is not None:

            if alias_used:

                match_type = "ALIAS"
                match_status = "MATCHED"

            else:

                match_type = "EXACT"
                match_status = "MATCHED"

            results.append({

                "source_index": index,

                "master_id": exact[
                    "master_id"
                ],

                "source_village":
                    original_village,

                "standard_village":
                    exact["village"],

                "source_cell":
                    original_cell,

                "standard_cell":
                    exact["cell"],

                "source_sector":
                    original_sector,

                "standard_sector":
                    exact["sector"],

                "district":
                    exact["district"],

                "match_type":
                    match_type,

                "match_score":
                    100,

                "match_status":
                    match_status,

                "alias_used":
                    alias_used,

                "duplicate_in_report":
                    row[
                        "duplicate_in_report"
                    ]
            })

            continue

        # ----------------------------------------------------
        # FUZZY MATCH WITHIN SAME CELL
        # ----------------------------------------------------

        cell_key = (
            district,
            sector,
            cell
        )

        candidates = cell_index.get(
            cell_key,
            []
        )

        matched_name, score = fuzzy_match(
            candidates,
            village
        )

        if matched_name is not None:

            candidates_df = master[
                (master["district_norm"] == district)
                &
                (master["sector_norm"] == sector)
                &
                (master["cell_norm"] == cell)
                &
                (master["village_norm"] == matched_name)
            ]

            if not candidates_df.empty:

                matched = candidates_df.iloc[0]

                score = round(
                    float(score),
                    2
                )

                if score >= AUTO_MATCH_SCORE:

                    status = "AUTO_FUZZY"

                elif score >= REVIEW_SCORE:

                    status = "REVIEW"

                else:

                    status = "UNMATCHED"

                results.append({

                    "source_index":
                        index,

                    "master_id":
                        matched["master_id"],

                    "source_village":
                        original_village,

                    "standard_village":
                        matched["village"],

                    "source_cell":
                        original_cell,

                    "standard_cell":
                        matched["cell"],

                    "source_sector":
                        original_sector,

                    "standard_sector":
                        matched["sector"],

                    "district":
                        matched["district"],

                    "match_type":
                        "FUZZY",

                    "match_score":
                        score,

                    "match_status":
                        status,

                    "alias_used":
                        alias_used,

                    "duplicate_in_report":
                        row[
                            "duplicate_in_report"
                        ]
                })

                continue

        # ----------------------------------------------------
        # UNMATCHED
        # ----------------------------------------------------

        results.append({

            "source_index":
                index,

            "master_id":
                None,

            "source_village":
                original_village,

            "standard_village":
                None,

            "source_cell":
                original_cell,

            "standard_cell":
                None,

            "source_sector":
                original_sector,

            "standard_sector":
                None,

            "district":
                original_district,

            "match_type":
                "NONE",

            "match_score":
                0,

            "match_status":
                "UNMATCHED",

            "alias_used":
                alias_used,

            "duplicate_in_report":
                row[
                    "duplicate_in_report"
                ]
        })

    match_df = pd.DataFrame(
        results
    )

    # --------------------------------------------------------
    # Merge matching results
    # --------------------------------------------------------

    df = df.reset_index()

    if "duplicate_in_report" in df.columns:

        df = df.drop(
            columns=[
                "duplicate_in_report"
            ]
        )

    df = df.merge(
        match_df,
        left_on="index",
        right_on="source_index",
        how="left"
    )

    # --------------------------------------------------------
    # Statistics
    # --------------------------------------------------------

    print("\nMATCHING RESULTS")

    print(
        f"Total records : {len(df):,}"
    )

    print(
        f"Exact / Alias : "
        f"{(df['match_status'] == 'MATCHED').sum():,}"
    )

    print(
        f"Auto fuzzy    : "
        f"{(df['match_status'] == 'AUTO_FUZZY').sum():,}"
    )

    print(
        f"Needs review  : "
        f"{(df['match_status'] == 'REVIEW').sum():,}"
    )

    print(
        f"Unmatched     : "
        f"{(df['match_status'] == 'UNMATCHED').sum():,}"
    )

    print(
        f"Duplicates    : "
        f"{df['duplicate_in_report'].sum():,}"
    )

    return df


# ============================================================
# CREATE STANDARDIZED DATA
# ============================================================

def create_standardized_data(df):

    result = df.copy()

    result["matched"] = (
        result["match_status"].isin(
            [
                "MATCHED",
                "AUTO_FUZZY"
            ]
        )
    )

    # --------------------------------------------------------
    # Replace geographic names with master names
    # --------------------------------------------------------

    for column in [
        "district",
        "sector",
        "cell",
        "village"
    ]:

        standard_column = (
            f"standard_{column}"
        )

        if standard_column in result.columns:

            mask = result["matched"]

            result.loc[
                mask,
                column
            ] = result.loc[
                mask,
                standard_column
            ]

    # --------------------------------------------------------
    # Numeric columns
    # --------------------------------------------------------

    numeric_columns = [

        "level1_households",
        "level1_members",
        "level1_amount",

        "level2_households",
        "level2_members",
        "level2_amount",

        "level3_households",
        "level3_members",
        "level3_amount",

        "level4_households",
        "level4_members",
        "level4_amount",

        "level5_households",
        "level5_members",
        "level5_amount"
    ]

    for column in numeric_columns:

        if column in result.columns:

            result[column] = pd.to_numeric(
                result[column],
                errors="coerce"
            ).fillna(0)

    # --------------------------------------------------------
    # TOTAL HOUSEHOLDS
    # --------------------------------------------------------

    result["total_households"] = (
        result["level1_households"]
        + result["level2_households"]
        + result["level3_households"]
        + result["level4_households"]
        + result["level5_households"]
    )

    # --------------------------------------------------------
    # TOTAL MEMBERS
    # --------------------------------------------------------

    result["total_members"] = (
        result["level1_members"]
        + result["level2_members"]
        + result["level3_members"]
        + result["level4_members"]
        + result["level5_members"]
    )

    # --------------------------------------------------------
    # TOTAL AMOUNT
    # --------------------------------------------------------

    result["total_amount"] = (
        result["level1_amount"]
        + result["level2_amount"]
        + result["level3_amount"]
        + result["level4_amount"]
        + result["level5_amount"]
    )

    # --------------------------------------------------------
    # Remove technical columns
    # --------------------------------------------------------

    technical_columns = [

        "district_norm",
        "sector_norm",
        "cell_norm",
        "village_norm",

        "source_index",
        "index",

        "standard_district",
        "standard_sector",
        "standard_cell",
        "standard_village",

        "matched"
    ]

    for column in technical_columns:

        if column in result.columns:

            result.drop(
                columns=[column],
                inplace=True
            )

    return result


# ============================================================
# CREATE MATCHING SUMMARY
# ============================================================

def create_matching_summary(
    all_matches
):

    summary = (
        all_matches
        .groupby(
            [
                "source_file",
                "match_status"
            ],
            dropna=False
        )
        .size()
        .reset_index(
            name="records"
        )
    )

    return summary


# ============================================================
# CREATE PENDING ALIASES
# ============================================================

def create_pending_aliases(
    all_matches,
    existing_aliases
):

    candidates = all_matches[
        all_matches["match_status"].isin(
            [
                "REVIEW",
                "UNMATCHED"
            ]
        )
    ].copy()

    if candidates.empty:
        return existing_aliases

    pending = []

    for _, row in candidates.iterrows():

        # Only create a candidate when a standard village exists
        if pd.isna(
            row.get("standard_village")
        ):
            continue

        standard_village = str(
            row["standard_village"]
        ).strip()

        if not standard_village:
            continue

        pending.append({

            "source_village":
                row["source_village"],

            "standard_village":
                standard_village,

            "source_cell":
                row["source_cell"],

            "standard_cell":
                row["standard_cell"],

            "source_sector":
                row["source_sector"],

            "standard_sector":
                row["standard_sector"],

            "status":
                "PENDING",

            "match_type":
                row["match_type"],

            "match_score":
                row["match_score"],

            "approved_by":
                "",

            "approved_date":
                ""
        })

    if not pending:
        return existing_aliases

    pending_df = pd.DataFrame(
        pending
    )

    combined = pd.concat(
        [
            existing_aliases,
            pending_df
        ],
        ignore_index=True
    )

    # Remove duplicates
    combined = combined.drop_duplicates(
        subset=[
            "source_village",
            "standard_village",
            "source_cell",
            "standard_cell",
            "source_sector",
            "standard_sector"
        ],
        keep="first"
    )

    return combined


# ============================================================
# MAIN
# ============================================================

def main():

    print("\n")
    print("=" * 80)
    print(
        "RUBAVU MUTUELLE VILLAGE MATCHING SYSTEM"
    )
    print("=" * 80)

    # --------------------------------------------------------
    # Check baseline
    # --------------------------------------------------------

    if not BASELINE_FILE.exists():

        raise FileNotFoundError(
            f"Baseline file not found: "
            f"{BASELINE_FILE}"
        )

    # --------------------------------------------------------
    # Read baseline
    # --------------------------------------------------------

    baseline = read_baseline()

    master = create_master_villages(
        baseline
    )

    master_lookup = create_master_lookup(
        master
    )

    cell_index = build_cell_index(
        master
    )

    # --------------------------------------------------------
    # Load permanent aliases
    # --------------------------------------------------------

    aliases = load_aliases()

    (
        village_aliases,
        cell_aliases,
        sector_aliases
    ) = build_alias_lookup(
        aliases
    )

    # --------------------------------------------------------
    # Find progress files
    # --------------------------------------------------------

    progress_files = sorted(
        [
            f
            for f in PROGRESS_DIR.glob(
                "RUBAVU_*.xlsx"
            )
            if not f.name.startswith("~$")
        ]
    )

    if not progress_files:

        raise FileNotFoundError(
            "No progress files found."
        )

    print("\nProgress files found:")

    for file in progress_files:

        print(
            f" - {file.name}"
        )

    # --------------------------------------------------------
    # Process all files
    # --------------------------------------------------------

    all_matches = []
    all_standardized = []

    for progress_file in progress_files:

        matched = match_progress_file(
            progress_file,
            master,
            cell_index,
            master_lookup,
            village_aliases,
            cell_aliases,
            sector_aliases
        )

        matched["source_file"] = (
            progress_file.name
        )

        all_matches.append(
            matched
        )

        # ----------------------------------------------------
        # Standardized data
        # ----------------------------------------------------

        standardized = create_standardized_data(
            matched
        )

        standardized["source_file"] = (
            progress_file.name
        )

        all_standardized.append(
            standardized
        )

        # ----------------------------------------------------
        # Save processed file
        # ----------------------------------------------------

        output_name = (
            progress_file.stem
            + "_matched.xlsx"
        )

        output_path = (
            PROCESSED_DIR
            / output_name
        )

        standardized.to_excel(
            output_path,
            index=False
        )

        print(
            "\nStandardized file created:"
        )

        print(
            output_path
        )

    # --------------------------------------------------------
    # Combine all matches
    # --------------------------------------------------------

    all_matches_df = pd.concat(
        all_matches,
        ignore_index=True
    )

    all_standardized_df = pd.concat(
        all_standardized,
        ignore_index=True
    )

    # --------------------------------------------------------
    # Matching summary
    # --------------------------------------------------------

    summary = create_matching_summary(
        all_matches_df
    )

    # --------------------------------------------------------
    # Review required
    # --------------------------------------------------------

    review = all_matches_df[
        all_matches_df["match_status"].isin(
            [
                "REVIEW",
                "UNMATCHED"
            ]
        )
    ].copy()

    # --------------------------------------------------------
    # Fuzzy matches
    # --------------------------------------------------------

    fuzzy = all_matches_df[
        all_matches_df["match_type"] == "FUZZY"
    ].copy()

    # --------------------------------------------------------
    # Duplicates
    # --------------------------------------------------------

    duplicates = all_matches_df[
        all_matches_df["duplicate_in_report"]
    ].copy()

    # ========================================================
    # UPDATE PERMANENT ALIAS FILE
    # ========================================================

    updated_aliases = create_pending_aliases(
        all_matches_df,
        aliases
    )

    save_aliases(
        updated_aliases
    )

    # ========================================================
    # SAVE MATCHING REPORT
    # ========================================================

    with pd.ExcelWriter(
        MATCH_REPORT,
        engine="openpyxl"
    ) as writer:

        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        all_matches_df.to_excel(
            writer,
            sheet_name="All Matches",
            index=False
        )

        review.to_excel(
            writer,
            sheet_name="Review Required",
            index=False
        )

        fuzzy.to_excel(
            writer,
            sheet_name="Fuzzy Matches",
            index=False
        )

        duplicates.to_excel(
            writer,
            sheet_name="Duplicates",
            index=False
        )

        master.to_excel(
            writer,
            sheet_name="Master Villages",
            index=False
        )

        all_standardized_df.to_excel(
            writer,
            sheet_name="Standardized Data",
            index=False
        )

        updated_aliases.to_excel(
            writer,
            sheet_name="Permanent Aliases",
            index=False
        )

    # ========================================================
    # CREATE VILLAGE NAME MAPPING
    # ========================================================

    mapping_columns = [

        "source_file",

        "source_village",
        "standard_village",

        "source_cell",
        "standard_cell",

        "source_sector",
        "standard_sector",

        "match_type",
        "match_score",
        "match_status",
        "alias_used"
    ]

    mapping_columns = [
        col
        for col in mapping_columns
        if col in all_matches_df.columns
    ]

    mapping = (
        all_matches_df[
            mapping_columns
        ]
        .drop_duplicates()
        .sort_values(
            by=[
                "source_sector",
                "source_cell",
                "source_village"
            ],
            na_position="last"
        )
    )

    mapping.to_excel(
        MAPPING_FILE,
        index=False
    )

    # ========================================================
    # FINAL REPORT
    # ========================================================

    print("\n")
    print("=" * 80)
    print("FINAL MATCHING RESULTS")
    print("=" * 80)

    print(
        f"Master villages:       "
        f"{len(master):,}"
    )

    print(
        f"Progress records:      "
        f"{len(all_matches_df):,}"
    )

    print(
        f"Exact / Alias matches: "
        f"{(all_matches_df['match_status'] == 'MATCHED').sum():,}"
    )

    print(
        f"Auto fuzzy matches:    "
        f"{(all_matches_df['match_status'] == 'AUTO_FUZZY').sum():,}"
    )

    print(
        f"Needs review:          "
        f"{(all_matches_df['match_status'] == 'REVIEW').sum():,}"
    )

    print(
        f"Unmatched:             "
        f"{(all_matches_df['match_status'] == 'UNMATCHED').sum():,}"
    )

    print(
        f"Duplicates:            "
        f"{all_matches_df['duplicate_in_report'].sum():,}"
    )

    approved_count = 0

    if not updated_aliases.empty:

        approved_count = (
            updated_aliases["status"]
            .astype(str)
            .str.upper()
            .eq("APPROVED")
            .sum()
        )

    pending_count = 0

    if not updated_aliases.empty:

        pending_count = (
            updated_aliases["status"]
            .astype(str)
            .str.upper()
            .eq("PENDING")
            .sum()
        )

    print(
        f"Permanent aliases:     "
        f"{approved_count:,}"
    )

    print(
        f"Pending aliases:       "
        f"{pending_count:,}"
    )

    # ========================================================
    # OUTPUT FILES
    # ========================================================

    print("\n")
    print("=" * 80)
    print("OUTPUT FILES")
    print("=" * 80)

    print(
        f"Matching report:"
    )

    print(
        f"  {MATCH_REPORT}"
    )

    print(
        f"\nPermanent aliases:"
    )

    print(
        f"  {ALIAS_FILE}"
    )

    print(
        f"\nName mapping:"
    )

    print(
        f"  {MAPPING_FILE}"
    )

    print(
        f"\nProcessed files:"
    )

    print(
        f"  {PROCESSED_DIR}"
    )

    print("\n")
    print("=" * 80)
    print("MATCHING COMPLETED")
    print("=" * 80)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()