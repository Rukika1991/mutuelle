import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
import re
import unicodedata


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Rubavu Mutuelle Monitoring",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

BASELINE_DIR = BASE_DIR / "data" / "baseline"
PROCESSED_DIR = BASE_DIR / "data" / "processed"

MASTER_FILE = BASELINE_DIR / "RubavunewAugust.xlsx"
MAPPING_FILE = BASE_DIR / "location_mapping.xlsx"


# ============================================================
# DATES
# ============================================================

CURRENT_DATE = pd.Timestamp("2026-08-22")
PREVIOUS_DATE = pd.Timestamp("2026-08-21")


# ============================================================
# EXPECTED OFFICIAL RUBAVU STRUCTURE
# ============================================================

EXPECTED_SECTORS = 12
EXPECTED_CELLS = 80
EXPECTED_VILLAGES = 519


# ============================================================
# CONTRIBUTION LEVELS
# ============================================================

LEVEL_CONTRIBUTIONS = {
    1: 0,
    2: 4000,
    3: 5000,
    4: 8000,
    5: 20000
}


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ===================================================== */
    /* GLOBAL APPLICATION */
    /* ===================================================== */

    .stApp {
        background-color: #f4f7fb;
        color: #111827;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }


    /* ===================================================== */
    /* HEADER */
    /* ===================================================== */

    .dashboard-header {
        background: linear-gradient(
            135deg,
            #111827,
            #1f2937
        );

        padding: 28px 32px;
        border-radius: 18px;

        margin-bottom: 25px;

        box-shadow:
            0 8px 24px
            rgba(0,0,0,0.15);
    }


    .dashboard-title {
        font-size: 32px;
        font-weight: 800;
        color: white;
        margin-bottom: 8px;
    }


    .dashboard-subtitle {
        font-size: 15px;
        color: #d1d5db;
        line-height: 1.6;
    }


    /* ===================================================== */
    /* SECTION TITLES */
    /* ===================================================== */

    .section-title {
        font-size: 22px;
        font-weight: 800;
        color: #111827;
        margin-top: 28px;
        margin-bottom: 14px;
    }


    /* ===================================================== */
    /* KPI CARDS */
    /* ===================================================== */

    .kpi-card {

        background: white;

        border: 1px solid #e5e7eb;

        border-radius: 16px;

        padding: 20px;

        min-height: 125px;

        box-shadow:
            0 4px 14px
            rgba(15,23,42,0.07);

        transition:
            transform 0.2s ease,
            box-shadow 0.2s ease;
    }


    .kpi-card:hover {

        transform: translateY(-3px);

        box-shadow:
            0 10px 25px
            rgba(15,23,42,0.12);
    }


    .kpi-label {

        font-size: 12px;

        font-weight: 700;

        color: #6b7280;

        text-transform: uppercase;

        letter-spacing: 0.5px;
    }


    .kpi-value {

        font-size: 25px;

        font-weight: 800;

        color: #111827;

        margin-top: 12px;

        word-break: break-word;
    }


    /* ===================================================== */
    /* STREAMLIT METRICS */
    /* ===================================================== */

    [data-testid="stMetric"] {

        background: white;

        border: 1px solid #e5e7eb;

        border-radius: 16px;

        padding: 18px;

        box-shadow:
            0 4px 14px
            rgba(15,23,42,0.07);
    }


    /* ===================================================== */
    /* DATAFRAME */
    /* ===================================================== */

    [data-testid="stDataFrame"] {

        background: white;

        border-radius: 14px;

        overflow: hidden;

        border: 1px solid #e5e7eb;
    }


    /* ===================================================== */
    /* EXPANDER */
    /* ===================================================== */

    [data-testid="stExpander"] {

        background: white;

        border: 1px solid #e5e7eb;

        border-radius: 14px;
    }


    /* ===================================================== */
    /* SIDEBAR */
    /* ===================================================== */

    [data-testid="stSidebar"] {

        background-color: #111827;
    }


    [data-testid="stSidebar"] * {
        color: white;
    }


    /* ===================================================== */
    /* DIVIDERS */
    /* ===================================================== */

    hr {

        border: none;

        border-top: 1px solid #d1d5db;

        margin-top: 30px;

        margin-bottom: 30px;
    }


    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# BASIC HELPERS
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


def clean_text(value):

    if pd.isna(value):
        return ""

    value = str(value)

    value = value.replace(
        "\u00a0",
        " "
    )

    value = value.strip()

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


def match_key(value):

    """
    Strong key used ONLY for matching.

    It ignores:

    - Capital/lowercase
    - Extra spaces
    - Hyphens
    - Underscores
    - Dots
    - Apostrophes
    - Accents
    """

    value = clean_text(value)

    value = unicodedata.normalize(
        "NFKD",
        value
    )

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


def to_number(series):

    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)


def money(value):

    return f"{value:,.0f} FRW"


def number(value):

    return f"{int(round(value)):,.0f}"


# ============================================================
# CSV DOWNLOAD
# ============================================================

def download_csv_button(df, filename, key):

    csv = df.to_csv(
        index=False
    ).encode(
        "utf-8-sig"
    )

    st.download_button(
        label="⬇️ Save as CSV",
        data=csv,
        file_name=filename,
        mime="text/csv",
        key=key
    )


# ============================================================
# MANUAL LOCATION CORRECTIONS
# ============================================================

MANUAL_LOCATION_MAPPING = {

    "sector": {},

    "cell": {

        # Kanyirabigog and Kanyirabigogo
        # are the same official cell
        "KANYIRABIGOG": "Kanyirabigogo",
        "KANYIRABIGOGO": "Kanyirabigogo",

        # Jenda and Jende are the same official cell
        "JENDA": "Jende",
        "JENDE": "Jende"
    },

    "village": {}
}


# ============================================================
# FIND PROGRESS FILES
# ============================================================

def find_progress_files():

    if not PROCESSED_DIR.exists():
        return []

    return sorted(
        PROCESSED_DIR.glob("*_matched.xlsx")
    )


# ============================================================
# EXTRACT DATE FROM FILE NAME
# ============================================================

def extract_date(filename):

    name = str(filename).upper()

    patterns = {

        "20_AUGUST": "2026-08-20",
        "21_AUGUST": "2026-08-21",
        "22_AUGUST": "2026-08-22",

        "20260820": "2026-08-20",
        "20260821": "2026-08-21",
        "20260822": "2026-08-22",

        "2026-08-20": "2026-08-20",
        "2026-08-21": "2026-08-21",
        "2026-08-22": "2026-08-22",
    }

    for pattern, date_value in patterns.items():

        if pattern in name:
            return pd.Timestamp(date_value)

    return pd.NaT


# ============================================================
# LOAD OFFICIAL MASTER BASELINE
# ============================================================

@st.cache_data
def load_master():

    if not MASTER_FILE.exists():
        return pd.DataFrame()

    df = pd.read_excel(
        MASTER_FILE
    )

    df = normalize_columns(df)

    for column in [
        "district",
        "sector",
        "cell",
        "village"
    ]:

        if column in df.columns:

            df[column] = (
                df[column]
                .apply(clean_text)
            )

    return df


# ============================================================
# LOAD LOCATION MAPPING
# ============================================================

@st.cache_data
def load_location_mapping():

    if not MAPPING_FILE.exists():
        return pd.DataFrame()

    try:

        mapping = pd.read_excel(
            MAPPING_FILE
        )

        mapping = normalize_columns(
            mapping
        )

        return mapping

    except Exception:

        return pd.DataFrame()


# ============================================================
# CREATE HIERARCHICAL MASTER LOOKUP
# ============================================================

def create_master_lookup(master):

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
# BUILD EXPLICIT CANONICAL MAPPING
# ============================================================

def build_explicit_mapping(mapping):

    mappings = {
        "sector": {},
        "cell": {},
        "village": {}
    }


    # --------------------------------------------------------
    # FIRST ADD MANUAL CORRECTIONS
    # --------------------------------------------------------

    for level, values in MANUAL_LOCATION_MAPPING.items():

        for original, canonical in values.items():

            mappings[level][
                match_key(original)
            ] = canonical


    # --------------------------------------------------------
    # THEN ADD EXCEL MAPPING
    # --------------------------------------------------------

    if mapping.empty:
        return mappings


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


    if (
        level_col is None
        or original_col is None
        or canonical_col is None
    ):

        return mappings


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
# APPLY LOCATION NORMALIZATION
# ============================================================

def apply_location_mapping(
    df,
    master_lookup,
    mapping
):

    df = df.copy()


    # --------------------------------------------------------
    # CLEAN ALL LOCATION COLUMNS
    # --------------------------------------------------------

    for column in [
        "district",
        "sector",
        "cell",
        "village"
    ]:

        if column in df.columns:

            df[column] = (
                df[column]
                .apply(clean_text)
            )


    explicit_mapping = build_explicit_mapping(
        mapping
    )


    # ========================================================
    # STEP 1: MAP SECTOR
    # ========================================================

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


        df["sector"] = df["sector"].apply(
            normalize_sector
        )


    # ========================================================
    # STEP 2: MAP CELL USING OFFICIAL SECTOR
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

            # Manual / explicit mapping first
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


        df["cell"] = df.apply(
            normalize_cell,
            axis=1
        )


    # ========================================================
    # STEP 3: MAP VILLAGE
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

            if village_key in explicit_mapping["village"]:

                village = explicit_mapping["village"][village_key]

            key = (
                match_key(sector),
                match_key(cell),
                match_key(village)
            )

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
# PREPARE BASELINE
# ============================================================

def prepare_baseline(df):

    df = df.copy()

    required = [
        "district",
        "sector",
        "cell",
        "village"
    ]

    missing = [
        column
        for column in required
        if column not in df.columns
    ]

    if missing:

        raise ValueError(
            "Baseline missing location columns: "
            + ", ".join(missing)
        )


    # --------------------------------------------------------
    # BASELINE MEMBERS
    # --------------------------------------------------------

    if "hhsize" in df.columns:

        df["baseline_members"] = to_number(
            df["hhsize"]
        )

    elif "household_size" in df.columns:

        df["baseline_members"] = to_number(
            df["household_size"]
        )

    else:

        raise ValueError(
            "Could not find hhsize or household_size."
        )


    # --------------------------------------------------------
    # BASELINE MONEY
    # --------------------------------------------------------

    if "total_amount" in df.columns:

        df["baseline_amount"] = to_number(
            df["total_amount"]
        )

    elif "cbhi_new_totalamount" in df.columns:

        df["baseline_amount"] = to_number(
            df["cbhi_new_totalamount"]
        )

    elif "amount" in df.columns:

        df["baseline_amount"] = to_number(
            df["amount"]
        )

    elif (
        "hhsize" in df.columns
        and "cbhi_level" in df.columns
    ):

        levels = (
            df["cbhi_level"]
            .astype(str)
            .str.extract(
                r"([1-5])",
                expand=False
            )
        )

        levels = pd.to_numeric(
            levels,
            errors="coerce"
        )

        contribution = (
            levels
            .map(LEVEL_CONTRIBUTIONS)
            .fillna(0)
        )

        df["baseline_amount"] = (
            df["baseline_members"]
            * contribution
        )

    else:

        raise ValueError(
            "Could not determine baseline amount."
        )


    return df


# ============================================================
# LOAD PROGRESS
# ============================================================

@st.cache_data
def load_progress():

    files = find_progress_files()

    if not files:
        return pd.DataFrame()


    frames = []


    for file in files:

        try:

            df = pd.read_excel(
                file
            )

            df = normalize_columns(
                df
            )

            df["source_file"] = file.name


            # ------------------------------------------------
            # DATE
            # ------------------------------------------------

            if "date" in df.columns:

                df["date"] = pd.to_datetime(
                    df["date"],
                    errors="coerce"
                )

            else:

                df["date"] = extract_date(
                    file.name
                )


            missing_date = (
                df["date"].isna()
            )

            if missing_date.any():

                extracted_date = extract_date(
                    file.name
                )

                df.loc[
                    missing_date,
                    "date"
                ] = extracted_date


            df["date"] = pd.to_datetime(
                df["date"],
                errors="coerce"
            ).dt.normalize()


            # ------------------------------------------------
            # LOCATIONS
            # ------------------------------------------------

            for column in [
                "district",
                "sector",
                "cell",
                "village"
            ]:

                if column in df.columns:

                    df[column] = (
                        df[column]
                        .apply(clean_text)
                    )


            frames.append(df)


        except Exception as e:

            st.warning(
                f"Could not read {file.name}: {e}"
            )


    if not frames:
        return pd.DataFrame()


    return pd.concat(
        frames,
        ignore_index=True
    )


# ============================================================
# PREPARE PROGRESS
# ============================================================

def prepare_progress(progress):

    progress = progress.copy()


    # --------------------------------------------------------
    # MONEY
    # --------------------------------------------------------

    amount_columns = [
        "level1_amount",
        "level2_amount",
        "level3_amount",
        "level4_amount",
        "level5_amount"
    ]


    for column in amount_columns:

        if column in progress.columns:

            progress[column] = to_number(
                progress[column]
            )


    if "total_amount" in progress.columns:

        progress["total_amount"] = to_number(
            progress["total_amount"]
        )

    else:

        existing = [
            column
            for column in amount_columns
            if column in progress.columns
        ]

        if existing:

            progress["total_amount"] = (
                progress[existing]
                .sum(axis=1)
            )

        else:

            progress["total_amount"] = 0


    # --------------------------------------------------------
    # MEMBERS
    # --------------------------------------------------------

    member_columns = [
        "level1_members",
        "level2_members",
        "level3_members",
        "level4_members",
        "level5_members"
    ]


    for column in member_columns:

        if column in progress.columns:

            progress[column] = to_number(
                progress[column]
            )


    if "total_members" in progress.columns:

        progress["total_members"] = to_number(
            progress["total_members"]
        )

    else:

        existing = [
            column
            for column in member_columns
            if column in progress.columns
        ]

        if existing:

            progress["total_members"] = (
                progress[existing]
                .sum(axis=1)
            )

        else:

            progress["total_members"] = 0


    return progress


# ============================================================
# PERFORMANCE TABLE
# ============================================================

def make_performance(
    baseline,
    current,
    previous,
    levels
):

    baseline_summary = (
        baseline
        .groupby(
            levels,
            dropna=False,
            as_index=False
        )
        .agg(
            Baseline_Money=(
                "baseline_amount",
                "sum"
            ),
            Baseline_Members=(
                "baseline_members",
                "sum"
            )
        )
    )


    # --------------------------------------------------------
    # CURRENT
    # --------------------------------------------------------

    if current.empty:

        current_summary = pd.DataFrame(
            columns=[
                *levels,
                "Total_Paid_22_Aug",
                "Total_Members_22_Aug"
            ]
        )

    else:

        current_summary = (
            current
            .groupby(
                levels,
                dropna=False,
                as_index=False
            )
            .agg(
                Total_Paid_22_Aug=(
                    "total_amount",
                    "sum"
                ),
                Total_Members_22_Aug=(
                    "total_members",
                    "sum"
                )
            )
        )


    # --------------------------------------------------------
    # PREVIOUS
    # --------------------------------------------------------

    if previous.empty:

        previous_summary = pd.DataFrame(
            columns=[
                *levels,
                "Previous_Paid_21_Aug",
                "Previous_Members_21_Aug"
            ]
        )

    else:

        previous_summary = (
            previous
            .groupby(
                levels,
                dropna=False,
                as_index=False
            )
            .agg(
                Previous_Paid_21_Aug=(
                    "total_amount",
                    "sum"
                ),
                Previous_Members_21_Aug=(
                    "total_members",
                    "sum"
                )
            )
        )


    # --------------------------------------------------------
    # MERGE
    # --------------------------------------------------------

    result = baseline_summary.merge(
        current_summary,
        on=levels,
        how="left"
    )

    result = result.merge(
        previous_summary,
        on=levels,
        how="left"
    )


    numeric_columns = [
        "Baseline_Money",
        "Baseline_Members",
        "Total_Paid_22_Aug",
        "Total_Members_22_Aug",
        "Previous_Paid_21_Aug",
        "Previous_Members_21_Aug"
    ]


    for column in numeric_columns:

        if column not in result.columns:

            result[column] = 0

        result[column] = (
            pd.to_numeric(
                result[column],
                errors="coerce"
            )
            .fillna(0)
        )


    # --------------------------------------------------------
    # TODAY
    # --------------------------------------------------------

    result["Today_Collected"] = (
        result["Total_Paid_22_Aug"]
        - result["Previous_Paid_21_Aug"]
    ).clip(lower=0)


    result["Today_Enrolled"] = (
        result["Total_Members_22_Aug"]
        - result["Previous_Members_21_Aug"]
    ).clip(lower=0)


    # --------------------------------------------------------
    # REMAINING
    # --------------------------------------------------------

    result["Remaining_Money"] = (
        result["Baseline_Money"]
        - result["Total_Paid_22_Aug"]
    ).clip(lower=0)


    result["Remaining_Members"] = (
        result["Baseline_Members"]
        - result["Total_Members_22_Aug"]
    ).clip(lower=0)


    # --------------------------------------------------------
    # PROGRESS
    # --------------------------------------------------------

    result["Money_Progress_%"] = (
        result["Total_Paid_22_Aug"]
        /
        result["Baseline_Money"].replace(
            0,
            pd.NA
        )
        * 100
    ).fillna(0)


    result["Member_Progress_%"] = (
        result["Total_Members_22_Aug"]
        /
        result["Baseline_Members"].replace(
            0,
            pd.NA
        )
        * 100
    ).fillna(0)


    # --------------------------------------------------------
    # RANK
    # --------------------------------------------------------

    result = (
        result
        .sort_values(
            [
                "Today_Collected",
                "Today_Enrolled"
            ],
            ascending=False
        )
        .reset_index(
            drop=True
        )
    )


    result.insert(
        0,
        "Rank",
        range(
            1,
            len(result) + 1
        )
    )


    return result


# ============================================================
# HIGHLIGHT ACTIVITY
# ============================================================

def highlight_activity(row):

    today_money = row.get(
        "Today_Collected",
        0
    )

    today_members = row.get(
        "Today_Enrolled",
        0
    )


    if (
        today_money == 0
        and today_members == 0
    ):

        return [
            (
                "background-color:#fee2e2;"
                "color:#991b1b;"
                "font-weight:700;"
            )
            for _ in row
        ]


    if today_money == 0:

        return [
            (
                "background-color:#fef3c7;"
                "color:#92400e;"
                "font-weight:600;"
            )
            for _ in row
        ]


    if today_members == 0:

        return [
            (
                "background-color:#dbeafe;"
                "color:#1e40af;"
                "font-weight:600;"
            )
            for _ in row
        ]


    return [
        ""
        for _ in row
    ]


# ============================================================
# TABLE FORMAT
# ============================================================

TABLE_FORMAT = {

    "Baseline_Money": "{:,.0f}",
    "Total_Paid_22_Aug": "{:,.0f}",
    "Previous_Paid_21_Aug": "{:,.0f}",
    "Today_Collected": "{:,.0f}",
    "Remaining_Money": "{:,.0f}",

    "Baseline_Members": "{:,.0f}",
    "Total_Members_22_Aug": "{:,.0f}",
    "Previous_Members_21_Aug": "{:,.0f}",
    "Today_Enrolled": "{:,.0f}",
    "Remaining_Members": "{:,.0f}",

    "Money_Progress_%": "{:.2f}%",
    "Member_Progress_%": "{:.2f}%"
}


# ============================================================
# LOAD DATA
# ============================================================

master = load_master()

mapping = load_location_mapping()

progress = load_progress()


# ============================================================
# VALIDATE FILES
# ============================================================

if master.empty:

    st.error(
        "❌ RubavunewAugust.xlsx was not found."
    )

    st.code(
        str(MASTER_FILE)
    )

    st.stop()


if progress.empty:

    st.error(
        "❌ No *_matched.xlsx progress files were found."
    )

    st.code(
        str(PROCESSED_DIR)
    )

    st.stop()


# ============================================================
# CREATE OFFICIAL MASTER LOOKUP
# ============================================================

master_lookup = create_master_lookup(
    master
)


# ============================================================
# NORMALIZE BASELINE
# ============================================================

master = apply_location_mapping(
    master,
    master_lookup,
    pd.DataFrame()
)


# ============================================================
# PREPARE BASELINE
# ============================================================

try:

    baseline = prepare_baseline(
        master
    )

except Exception as e:

    st.error(
        f"❌ Baseline preparation error: {e}"
    )

    st.stop()


# ============================================================
# STRICT STRUCTURE VALIDATION
# ============================================================

official_sectors = (
    baseline["sector"]
    .nunique()
)


official_cells = (
    baseline[
        ["sector", "cell"]
    ]
    .drop_duplicates()
    .shape[0]
)


official_villages = (
    baseline[
        ["sector", "cell", "village"]
    ]
    .drop_duplicates()
    .shape[0]
)


structure_valid = True


if official_sectors != EXPECTED_SECTORS:

    st.error(
        f"❌ Expected {EXPECTED_SECTORS} sectors "
        f"but baseline contains {official_sectors}."
    )

    structure_valid = False


if official_cells != EXPECTED_CELLS:

    st.error(
        f"❌ Expected {EXPECTED_CELLS} cells "
        f"but baseline contains {official_cells}."
    )

    structure_valid = False


if official_villages != EXPECTED_VILLAGES:

    st.error(
        f"❌ Expected {EXPECTED_VILLAGES} villages "
        f"but baseline contains {official_villages}."
    )

    structure_valid = False


if not structure_valid:

    st.warning(
        "The dashboard stopped because the official "
        "baseline structure does not match "
        "12 sectors / 80 cells / 519 villages."
    )

    st.stop()


# ============================================================
# NORMALIZE PROGRESS AGAINST OFFICIAL BASELINE
# ============================================================

progress = apply_location_mapping(
    progress,
    master_lookup,
    mapping
)


# ============================================================
# PREPARE PROGRESS
# ============================================================

progress = prepare_progress(
    progress
)


# ============================================================
# CURRENT / PREVIOUS DATA
# ============================================================

current = progress[
    progress["date"] == CURRENT_DATE
].copy()


previous = progress[
    progress["date"] == PREVIOUS_DATE
].copy()


if current.empty:

    st.error(
        "❌ No progress data found for 22/08/2026."
    )

    available_dates = (
        progress["date"]
        .dropna()
        .dt.strftime("%Y-%m-%d")
        .unique()
    )

    st.write(
        "Available dates:"
    )

    st.write(
        sorted(available_dates)
    )

    st.stop()


# ============================================================
# OVERALL BASELINE
# ============================================================

baseline_money = (
    baseline["baseline_amount"]
    .sum()
)

baseline_members = (
    baseline["baseline_members"]
    .sum()
)


# ============================================================
# OVERALL CURRENT
# ============================================================

total_paid = (
    current["total_amount"]
    .sum()
)

total_members = (
    current["total_members"]
    .sum()
)


# ============================================================
# OVERALL PREVIOUS
# ============================================================

previous_paid = (
    previous["total_amount"].sum()
    if not previous.empty
    else 0
)

previous_members = (
    previous["total_members"].sum()
    if not previous.empty
    else 0
)


# ============================================================
# TODAY
# ============================================================

today_collected = max(
    0,
    total_paid - previous_paid
)

today_enrolled = max(
    0,
    total_members - previous_members
)


# ============================================================
# REMAINING
# ============================================================

remaining_money = max(
    0,
    baseline_money - total_paid
)

remaining_members = max(
    0,
    baseline_members - total_members
)


# ============================================================
# PROGRESS %
# ============================================================

money_progress = (
    total_paid
    / baseline_money
    * 100
    if baseline_money > 0
    else 0
)

member_progress = (
    total_members
    / baseline_members
    * 100
    if baseline_members > 0
    else 0
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="dashboard-header">

        <div class="dashboard-title">
            📊 Rubavu Mutuelle Monitoring Dashboard
        </div>

        <div class="dashboard-subtitle">
            Official Baseline Monitoring • Progress as of 22/08/2026
            <br>
            Today Performance = 22/08/2026 − 21/08/2026
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# DATA VALIDATION
# ============================================================

with st.expander(
    "ℹ️ Data Sources & Validation",
    expanded=False
):

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.success(
            f"✅ {official_sectors} Official Sectors"
        )

    with col2:
        st.success(
            f"✅ {official_cells} Official Cells"
        )

    with col3:
        st.success(
            f"✅ {official_villages} Official Villages"
        )

    with col4:
        st.info(
            f"Baseline Records: {len(baseline):,}"
        )

    st.divider()

    st.markdown(
        f"""
        <div style="
            color:#000000 !important;
            font-size:15px;
            line-height:1.9;
        ">
            <b>Official baseline:</b> {MASTER_FILE.name}<br>
            <b>Location mapping:</b> {MAPPING_FILE.name}<br>
            <b>22/08/2026 progress records:</b> {len(current):,}<br>
            <b>21/08/2026 progress records:</b> {len(previous):,}
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# MONEY PERFORMANCE
# ============================================================

st.markdown(
    '<div class="section-title">💰 Money Performance</div>',
    unsafe_allow_html=True
)


c1, c2, c3, c4, c5 = st.columns(5)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Baseline</div>
            <div class="kpi-value">
                {money(baseline_money)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Paid — 22/08</div>
            <div class="kpi-value">
                {money(total_paid)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Previous Paid — 21/08</div>
            <div class="kpi-value">
                {money(previous_paid)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Today Collected</div>
            <div class="kpi-value">
                {money(today_collected)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c5:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Remaining</div>
            <div class="kpi-value">
                {money(remaining_money)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# MEMBER PERFORMANCE
# ============================================================

st.markdown(
    '<div class="section-title">👥 Member Enrollment</div>',
    unsafe_allow_html=True
)


c1, c2, c3, c4, c5 = st.columns(5)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Baseline Members</div>
            <div class="kpi-value">
                {number(baseline_members)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Total — 22/08</div>
            <div class="kpi-value">
                {number(total_members)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Previous — 21/08</div>
            <div class="kpi-value">
                {number(previous_members)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Today Enrolled</div>
            <div class="kpi-value">
                {number(today_enrolled)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c5:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Remaining Members</div>
            <div class="kpi-value">
                {number(remaining_members)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# OVERALL PROGRESS
# ============================================================

st.markdown(
    '<div class="section-title">📈 Overall Progress</div>',
    unsafe_allow_html=True
)


c1, c2 = st.columns(2)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                💰 Money Progress
            </div>

            <div class="kpi-value"
                 style="color:#000000 !important;">
                {money_progress:.2f}%
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                👥 Member Progress
            </div>

            <div class="kpi-value"
                 style="color:#000000 !important;">
                {member_progress:.2f}%
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# SECTOR PERFORMANCE
# ============================================================

sector = make_performance(
    baseline,
    current,
    previous,
    ["sector"]
)


# ============================================================
# SECTOR RANKING
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">🥇 Sector Ranking</div>',
    unsafe_allow_html=True
)


st.dataframe(

    sector.style
    .apply(
        highlight_activity,
        axis=1
    )
    .format(
        TABLE_FORMAT
    ),

    use_container_width=True,

    hide_index=True,

    height=550
)


download_csv_button(
    sector,
    "Rubavu_Sector_Ranking_22-08-2026.csv",
    "download_sector_ranking"
)


# ============================================================
# ZERO-ACTIVITY SECTORS
# ============================================================

zero_sector = sector[
    (
        sector["Today_Collected"] == 0
    )
    |
    (
        sector["Today_Enrolled"] == 0
    )
]


if not zero_sector.empty:

    st.warning(
        f"⚠️ {len(zero_sector)} sector(s) have "
        "zero money collection or zero enrollment today."
    )

    download_csv_button(
        zero_sector,
        "Rubavu_Zero_Activity_Sectors_22-08-2026.csv",
        "download_zero_sectors"
    )


# ============================================================
# CELL PERFORMANCE
# ============================================================

cell = make_performance(
    baseline,
    current,
    previous,
    [
        "sector",
        "cell"
    ]
)


# ============================================================
# CELL RANKING
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">📍 Cell Ranking</div>',
    unsafe_allow_html=True
)


st.dataframe(

    cell.style
    .apply(
        highlight_activity,
        axis=1
    )
    .format(
        TABLE_FORMAT
    ),

    use_container_width=True,

    hide_index=True,

    height=650
)


download_csv_button(
    cell,
    "Rubavu_Cell_Ranking_22-08-2026.csv",
    "download_cell_ranking"
)


# ============================================================
# ZERO-ACTIVITY CELLS
# ============================================================

st.markdown(
    "### 🚨 Cells With Zero Activity Today"
)


zero_cells = cell[
    (
        cell["Today_Collected"] == 0
    )
    |
    (
        cell["Today_Enrolled"] == 0
    )
]


if zero_cells.empty:

    st.success(
        "All cells have money collection "
        "and enrollment activity today."
    )

else:

    st.dataframe(

        zero_cells.style
        .apply(
            highlight_activity,
            axis=1
        )
        .format(
            TABLE_FORMAT
        ),

        use_container_width=True,

        hide_index=True,

        height=500
    )


    download_csv_button(
        zero_cells,
        "Rubavu_Zero_Activity_Cells_22-08-2026.csv",
        "download_zero_cells"
    )


# ============================================================
# VILLAGE PERFORMANCE
# ============================================================

village = make_performance(
    baseline,
    current,
    previous,
    [
        "sector",
        "cell",
        "village"
    ]
)


# ============================================================
# VILLAGE RANKING
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">🏘️ Village Ranking</div>',
    unsafe_allow_html=True
)


st.dataframe(

    village.style
    .apply(
        highlight_activity,
        axis=1
    )
    .format(
        TABLE_FORMAT
    ),

    use_container_width=True,

    hide_index=True,

    height=750
)


download_csv_button(
    village,
    "Rubavu_Village_Ranking_22-08-2026.csv",
    "download_village_ranking"
)


# ============================================================
# ZERO-ACTIVITY VILLAGES
# ============================================================

st.markdown(
    "### 🚨 Villages With Zero Activity Today"
)


zero_villages = village[
    (
        village["Today_Collected"] == 0
    )
    |
    (
        village["Today_Enrolled"] == 0
    )
]


if zero_villages.empty:

    st.success(
        "All villages have money collection "
        "and enrollment activity today."
    )

else:

    st.dataframe(

        zero_villages.style
        .apply(
            highlight_activity,
            axis=1
        )
        .format(
            TABLE_FORMAT
        ),

        use_container_width=True,

        hide_index=True,

        height=550
    )


    download_csv_button(
        zero_villages,
        "Rubavu_Zero_Activity_Villages_22-08-2026.csv",
        "download_zero_villages"
    )


# ============================================================
# TOP 30 VILLAGES
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">🏆 Top 30 Villages — Today</div>',
    unsafe_allow_html=True
)


top30 = village.head(30).copy()


fig = px.bar(

    top30,

    x="village",

    y="Today_Collected",

    color="sector",

    text="Today_Collected",

    title="Today's Collection — Top 30 Villages"
)


fig.update_traces(

    texttemplate="%{text:,.0f}",

    textposition="outside"
)


fig.update_layout(

    xaxis_title="Village",

    yaxis_title="Today's Collection (FRW)",

    xaxis_tickangle=-45,

    height=600
)


st.plotly_chart(

    fig,

    use_container_width=True
)


# ============================================================
# CSV DOWNLOAD - TOP 30 VILLAGES
# ============================================================

download_csv_button(
    top30,
    "Rubavu_Top_30_Villages_22-08-2026.csv",
    "download_top30_villages"
)


# ============================================================
# FOOTER
# ============================================================

st.divider()


st.caption(
    "Rubavu District Mutuelle Monitoring | "
    "Official Structure: 12 Sectors • 80 Cells • 519 Villages | "
    "Official baseline: RubavunewAugust.xlsx | "
    "Today = 22/08/2026 − 21/08/2026"
)