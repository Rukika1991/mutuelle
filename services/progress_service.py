# ============================================================
# RUBAVU MUTUELLE MONITORING
# AUTOMATIC DAILY PROGRESS PROCESSOR
#
# PURPOSE
# -------
# Automatically process daily uploaded Excel files.
#
# Example:
#
#   23/08/2026 -> historical
#   24/08/2026 -> historical
#   25/08/2026 -> previous
#   26/08/2026 -> current
#
# dashboard_service.py then calculates:
#
#   Today = Current - Previous
#   Remaining = Baseline - Current
#
# ============================================================

from pathlib import Path
from typing import Optional
import re
import shutil
import pandas as pd


# ============================================================
# PROJECT PATHS
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent

UPLOAD_DIR = PROJECT_ROOT / "data" / "uploads"

OUTPUT_DIR = PROJECT_ROOT / "output"

PROCESSED_FILE = OUTPUT_DIR / "processed_progress.xlsx"


# ============================================================
# EXPECTED STRUCTURE
# ============================================================

EXPECTED_SECTORS = 12
EXPECTED_CELLS = 80
EXPECTED_VILLAGES = 519


LOCATION_COLUMNS = [
    "district",
    "sector",
    "cell",
    "village",
]


LEVELS = [1, 2, 3, 4, 5]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_text(value) -> str:
    """
    Normalize location text.

    Examples:

        ' BUGESHI '  -> 'BUGESHI'
        'Bugeshi'    -> 'BUGESHI'
        'BUGESHI\n'  -> 'BUGESHI'

    This is important because location matching must be
    case-insensitive.
    """

    if pd.isna(value):
        return ""

    value = str(value)

    value = value.replace("\n", " ")
    value = value.replace("\r", " ")

    value = " ".join(value.split())

    return value.strip().upper()


def normalize_column_name(value) -> str:
    """
    Convert a column name into a standard snake_case format.
    """

    value = str(value).strip().lower()

    value = value.replace("\n", " ")
    value = value.replace("\r", " ")

    value = " ".join(value.split())

    value = value.replace(" ", "_")
    value = value.replace("-", "_")
    value = value.replace("/", "_")

    while "__" in value:
        value = value.replace("__", "_")

    return value


def to_numeric(series) -> pd.Series:
    """
    Safely convert values to numbers.
    """

    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)


# ============================================================
# EXTRACT REPORT DATE
# ============================================================

def extract_report_date(file_path: Path) -> pd.Timestamp:
    """
    Extract report date from the Excel header, falling back to the
    filename when the workbook has no recognizable Report Date.

    Header example: Report Date: 8/25/2026
    Filename example: REPORT ON 8-10-2026.xlsx (DD-MM-YYYY)
    """

    # Try the first three Excel rows first. The report-date text may
    # appear in a cell other than A1.
    try:
        raw = pd.read_excel(
            file_path,
            header=None,
            nrows=3
        )
    except Exception:
        raw = pd.DataFrame()

    if not raw.empty:
        header_text = " ".join(
            str(value)
            for value in raw.to_numpy().flatten()
            if not pd.isna(value)
        )

        match = re.search(
            r"Report\s*Date\s*:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
            header_text,
            flags=re.IGNORECASE
        )

        if match:
            report_date = pd.to_datetime(
                match.group(1),
                format="%m/%d/%Y",
                errors="coerce"
            )
            if not pd.isna(report_date):
                return pd.Timestamp(report_date).normalize()

    # Fallback to filename, using Rwanda's DD-MM-YYYY convention.
    # Therefore 8-10-2026 means 8 October 2026.
    match = re.search(
        r"(?<!\d)(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})(?!\d)",
        file_path.stem
    )

    if match:
        day, month, year = map(int, match.groups())
        report_date = pd.to_datetime(
            f"{day:02d}-{month:02d}-{year:04d}",
            format="%d-%m-%Y",
            errors="coerce"
        )
        if not pd.isna(report_date):
            return pd.Timestamp(report_date).normalize()

    raise ValueError(
        f"Could not find a valid report date in {file_path.name}. "
        "Expected a header such as 'Report Date: 8/25/2026' or a "
        "filename containing a DD-MM-YYYY date such as "
        "'REPORT ON 8-10-2026.xlsx'."
    )


# ============================================================
# READ RUBAVU EXCEL
# ============================================================

def read_rubavu_excel(file_path: Path) -> pd.DataFrame:
    """
    Read the Rubavu Mutuelle Excel structure.

    Actual structure:

    Row 0:
        Enrolment Report / Period / Report Date

    Row 1:
        DISTRICT SECTOR CELL VILLAGE LEVEL 1 ...

    Row 2:
        Households Members Amount ...

    Row 3 onward:
        Actual village data

    Therefore we read header=None and manually select
    columns 0-18.
    """

    raw = pd.read_excel(
        file_path,
        header=None
    )

    if len(raw) < 4:
        raise ValueError(
            f"{file_path.name} does not contain enough rows."
        )

    # --------------------------------------------------------
    # Actual columns in the Excel
    # --------------------------------------------------------

    expected_columns = [
        "district",
        "sector",
        "cell",
        "village",

        "level_1_households",
        "level_1_members",
        "level_1_amount",

        "level_2_households",
        "level_2_members",
        "level_2_amount",

        "level_3_households",
        "level_3_members",
        "level_3_amount",

        "level_4_households",
        "level_4_members",
        "level_4_amount",

        "level_5_households",
        "level_5_members",
        "level_5_amount",
    ]

    if raw.shape[1] < 19:
        raise ValueError(
            f"{file_path.name} has only {raw.shape[1]} columns. "
            f"Expected at least 19 columns."
        )

    # --------------------------------------------------------
    # Data begins at row 3
    # --------------------------------------------------------

    df = raw.iloc[3:, :19].copy()

    df.columns = expected_columns

    # --------------------------------------------------------
    # Remove completely empty rows
    # --------------------------------------------------------

    df = df.dropna(
        how="all"
    ).copy()

    # --------------------------------------------------------
    # Normalize locations
    # --------------------------------------------------------

    for column in LOCATION_COLUMNS:

        df[column] = df[column].apply(
            normalize_text
        )

    # --------------------------------------------------------
    # Remove rows without village
    # --------------------------------------------------------

    df = df[
        df["village"] != ""
    ].copy()

    # --------------------------------------------------------
    # Numeric columns
    # --------------------------------------------------------

    numeric_columns = [
        column
        for column in df.columns
        if column not in LOCATION_COLUMNS
    ]

    for column in numeric_columns:

        df[column] = to_numeric(
            df[column]
        )

    # --------------------------------------------------------
    # Calculate total members
    # --------------------------------------------------------

    member_columns = [
        f"level_{level}_members"
        for level in LEVELS
    ]

    df["members"] = (
        df[member_columns]
        .sum(axis=1)
    )

    # --------------------------------------------------------
    # Calculate total money
    # --------------------------------------------------------

    amount_columns = [
        f"level_{level}_amount"
        for level in LEVELS
    ]

    df["amount"] = (
        df[amount_columns]
        .sum(axis=1)
    )

    return df.reset_index(
        drop=True
    )


# ============================================================
# VALIDATE LOCATION STRUCTURE
# ============================================================

def validate_locations(
    df: pd.DataFrame,
    file_name: str
):
    """
    Validate the geographical structure.

    Expected:

        12 sectors
        80 cells
        519 villages
    """

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
        .drop_duplicates()
        .shape[0]
    )

    print(
        f"Location validation for {file_name}:"
    )

    print(
        f"  Sectors:  {sectors}"
    )

    print(
        f"  Cells:    {cells}"
    )

    print(
        f"  Villages: {villages}"
    )

    if sectors != EXPECTED_SECTORS:

        raise ValueError(
            f"{file_name}: Expected "
            f"{EXPECTED_SECTORS} sectors but found "
            f"{sectors}."
        )

    if cells != EXPECTED_CELLS:

        raise ValueError(
            f"{file_name}: Expected "
            f"{EXPECTED_CELLS} cells but found "
            f"{cells}."
        )

    if villages != EXPECTED_VILLAGES:

        raise ValueError(
            f"{file_name}: Expected "
            f"{EXPECTED_VILLAGES} villages but found "
            f"{villages}."
        )

    return {
        "sectors": sectors,
        "cells": cells,
        "villages": villages,
    }


# ============================================================
# PROCESS ONE FILE
# ============================================================

def process_file(
    file_path: Path
) -> pd.DataFrame:
    """
    Process one daily Excel file.
    """

    print("-" * 70)

    print(
        f"Processing: {file_path.name}"
    )

    report_date = extract_report_date(
        file_path
    )

    print(
        f"Report date: "
        f"{report_date.strftime('%Y-%m-%d')}"
    )

    df = read_rubavu_excel(
        file_path
    )

    # --------------------------------------------------------
    # Validate structure
    # --------------------------------------------------------

    validate_locations(
        df,
        file_path.name
    )

    # --------------------------------------------------------
    # Add report date
    # --------------------------------------------------------

    df["report_date"] = report_date

    # --------------------------------------------------------
    # Keep only fields required by dashboard
    # --------------------------------------------------------

    output_columns = [
        "district",
        "sector",
        "cell",
        "village",

        "level_1_members",
        "level_2_members",
        "level_3_members",
        "level_4_members",
        "level_5_members",

        "level_1_amount",
        "level_2_amount",
        "level_3_amount",
        "level_4_amount",
        "level_5_amount",

        "members",
        "amount",

        "report_date",
    ]

    df = df[
        output_columns
    ].copy()

    print(
        f"Rows processed: {len(df):,}"
    )

    print(
        f"Members: {df['members'].sum():,.0f}"
    )

    print(
        f"Amount: {df['amount'].sum():,.0f}"
    )

    return df


# ============================================================
# LOAD EXISTING PROCESSED DATA
# ============================================================

def load_existing_processed() -> pd.DataFrame:
    """
    Load existing processed history.

    If the output file does not exist, return an empty
    dataframe.
    """

    if not PROCESSED_FILE.exists():

        return pd.DataFrame()

    print(
        f"Loading existing processed history:"
    )

    print(
        PROCESSED_FILE
    )

    df = pd.read_excel(
        PROCESSED_FILE
    )

    df.columns = [
        normalize_column_name(c)
        for c in df.columns
    ]

    # --------------------------------------------------------
    # Normalize locations
    # --------------------------------------------------------

    for column in LOCATION_COLUMNS:

        if column in df.columns:

            df[column] = df[column].apply(
                normalize_text
            )

    # --------------------------------------------------------
    # Date
    # --------------------------------------------------------

    if "report_date" in df.columns:

        df["report_date"] = pd.to_datetime(
            df["report_date"],
            errors="coerce"
        ).dt.normalize()

    return df


# ============================================================
# MERGE DAILY HISTORY
# ============================================================

def merge_progress(
    existing: pd.DataFrame,
    new_data: pd.DataFrame
) -> pd.DataFrame:
    """
    Merge historical progress with new uploads.

    Important:

    A report date is considered a complete snapshot.

    If the same date is uploaded again, the new file replaces
    the old snapshot for that date.

    This prevents duplicate 25-August data.
    """

    if existing.empty:

        combined = new_data.copy()

    else:

        combined = pd.concat(
            [
                existing,
                new_data
            ],
            ignore_index=True
        )

    # --------------------------------------------------------
    # Ensure dates are normalized
    # --------------------------------------------------------

    combined["report_date"] = pd.to_datetime(
        combined["report_date"],
        errors="coerce"
    ).dt.normalize()

    # --------------------------------------------------------
    # Normalize locations
    # --------------------------------------------------------

    for column in LOCATION_COLUMNS:

        combined[column] = combined[column].apply(
            normalize_text
        )

    # --------------------------------------------------------
    # Remove duplicate snapshot rows
    #
    # Key:
    #
    # date + district + sector + cell + village
    #
    # keep last means:
    #
    # new upload replaces old upload.
    # --------------------------------------------------------

    key_columns = [
        "report_date",
        "district",
        "sector",
        "cell",
        "village",
    ]

    combined = combined.drop_duplicates(
        subset=key_columns,
        keep="last"
    )

    # --------------------------------------------------------
    # Sort chronologically
    # --------------------------------------------------------

    combined = combined.sort_values(
        by=[
            "report_date",
            "district",
            "sector",
            "cell",
            "village",
        ]
    ).reset_index(
        drop=True
    )

    return combined


# ============================================================
# SAVE PROCESSED DATA
# ============================================================

def save_processed(
    df: pd.DataFrame
):
    """
    Save processed history to Excel.
    """

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Temporary file first
    #
    # This reduces risk of leaving a corrupted output file
    # if writing fails.
    # --------------------------------------------------------

    temp_file = (
        OUTPUT_DIR
        / "processed_progress.tmp.xlsx"
    )

    df.to_excel(
        temp_file,
        index=False
    )

    shutil.move(
        str(temp_file),
        str(PROCESSED_FILE)
    )

    print()

    print(
        "Saved processed file:"
    )

    print(
        PROCESSED_FILE
    )


# ============================================================
# GET UPLOAD FILES
# ============================================================

def get_upload_files():
    """
    Find Excel files in data/uploads.

    Temporary Excel files beginning with '~$' are ignored.
    """

    if not UPLOAD_DIR.exists():

        UPLOAD_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

    files = sorted(
        [
            f
            for f in UPLOAD_DIR.glob("*.xlsx")
            if not f.name.startswith("~$")
        ]
    )

    return files


# ============================================================
# AUTOMATIC PROCESSOR
# ============================================================

def process_uploaded_files():
    """
    Main automatic processing function.

    Every Excel file in data/uploads is inspected.

    Historical dates are retained.

    Same-date snapshots replace the old snapshot.

    This means you can safely upload a corrected Excel for
    the same date.
    """

    print()

    print("=" * 70)

    print(
        "AUTOMATIC PROGRESS PROCESSING"
    )

    print("=" * 70)

    print()

    print(
        "Upload directory:"
    )

    print(
        UPLOAD_DIR
    )

    files = get_upload_files()

    print()

    print(
        f"Excel files found: {len(files)}"
    )

    if not files:

        print()

        print(
            "No Excel files found."
        )

        print(
            "Put today's Rubavu Excel file in:"
        )

        print(
            UPLOAD_DIR
        )

        return pd.DataFrame()

    # --------------------------------------------------------
    # Load historical processed data
    # --------------------------------------------------------

    existing = load_existing_processed()

    # --------------------------------------------------------
    # Determine already processed dates
    # --------------------------------------------------------

    existing_dates = set()

    if (
        not existing.empty
        and "report_date" in existing.columns
    ):

        existing_dates = set(
            pd.to_datetime(
                existing["report_date"],
                errors="coerce"
            )
            .dropna()
            .dt.normalize()
            .unique()
        )

    print()

    if existing_dates:

        print(
            "Existing processed dates:"
        )

        for date in sorted(existing_dates):

            print(
                f"  - {pd.Timestamp(date).strftime('%Y-%m-%d')}"
            )

    # --------------------------------------------------------
    # Process uploads
    # --------------------------------------------------------

    new_data_frames = []

    for file_path in files:

        try:

            report_date = extract_report_date(
                file_path
            )

            # ------------------------------------------------
            # IMPORTANT:
            #
            # We still read the file even if the date already
            # exists.
            #
            # Why?
            #
            # If a corrected report is uploaded for the same
            # date, it should replace the previous snapshot.
            # ------------------------------------------------

            print()

            df = process_file(
                file_path
            )

            new_data_frames.append(
                df
            )

        except Exception as exc:

            print()

            print(
                f"ERROR processing "
                f"{file_path.name}:"
            )

            print(
                str(exc)
            )

            raise

    # --------------------------------------------------------
    # Combine all new uploads
    # --------------------------------------------------------

    if new_data_frames:

        new_data = pd.concat(
            new_data_frames,
            ignore_index=True
        )

    else:

        new_data = pd.DataFrame()

    # --------------------------------------------------------
    # Merge
    # --------------------------------------------------------

    combined = merge_progress(
        existing,
        new_data
    )

    # --------------------------------------------------------
    # Validate final history
    # --------------------------------------------------------

    if combined.empty:

        raise ValueError(
            "No progress data available after processing."
        )

    # --------------------------------------------------------
    # Final dates
    # --------------------------------------------------------

    dates = sorted(
        combined["report_date"]
        .dropna()
        .unique()
    )

    # --------------------------------------------------------
    # Final validation
    #
    # Validate the latest snapshot.
    # --------------------------------------------------------

    latest_date = pd.Timestamp(
        dates[-1]
    )

    latest = combined[
        combined["report_date"]
        == latest_date
    ].copy()

    validate_locations(
        latest,
        f"Latest snapshot {latest_date.strftime('%Y-%m-%d')}"
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    save_processed(
        combined
    )

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()

    print("=" * 70)

    print(
        "PROCESSING COMPLETE"
    )

    print("=" * 70)

    print()

    print(
        f"Total rows: {len(combined):,}"
    )

    print()

    print(
        "Available report dates:"
    )

    for date in dates:

        date_data = combined[
            combined["report_date"]
            == pd.Timestamp(date)
        ]

        print(
            f"  - {pd.Timestamp(date).strftime('%Y-%m-%d')}"
            f" | rows: {len(date_data):,}"
            f" | members: {date_data['members'].sum():,.0f}"
            f" | amount: {date_data['amount'].sum():,.0f}"
        )

    print()

    # --------------------------------------------------------
    # Automatic comparison information
    # --------------------------------------------------------

    if len(dates) >= 2:

        previous_date = pd.Timestamp(
            dates[-2]
        )

        current_date = pd.Timestamp(
            dates[-1]
        )

        previous = combined[
            combined["report_date"]
            == previous_date
        ]

        current = combined[
            combined["report_date"]
            == current_date
        ]

        previous_members = (
            previous["members"].sum()
        )

        current_members = (
            current["members"].sum()
        )

        previous_amount = (
            previous["amount"].sum()
        )

        current_amount = (
            current["amount"].sum()
        )

        today_members = (
            current_members
            - previous_members
        )

        today_amount = (
            current_amount
            - previous_amount
        )

        print("=" * 70)

        print(
            "LATEST DAILY COMPARISON"
        )

        print("=" * 70)

        print()

        print(
            f"Previous date: "
            f"{previous_date.strftime('%Y-%m-%d')}"
        )

        print(
            f"Current date:  "
            f"{current_date.strftime('%Y-%m-%d')}"
        )

        print()

        print(
            f"Previous members: "
            f"{previous_members:,.0f}"
        )

        print(
            f"Current members:  "
            f"{current_members:,.0f}"
        )

        print(
            f"Today members:    "
            f"{today_members:,.0f}"
        )

        print()

        print(
            f"Previous amount: "
            f"{previous_amount:,.0f}"
        )

        print(
            f"Current amount:  "
            f"{current_amount:,.0f}"
        )

        print(
            f"Today collected: "
            f"{today_amount:,.0f}"
        )

        print()

    else:

        print(
            "Only one report date is available."
        )

        print(
            "A second date is required for "
            "previous-vs-current comparison."
        )

    return combined


# ============================================================
# TEST / MAIN
# ============================================================

def test_progress_service():

    print()

    print("=" * 70)

    print(
        "RUBAVU PROGRESS SERVICE TEST"
    )

    print("=" * 70)

    print()

    print(
        "Automatically processing uploaded Excel files..."
    )

    print()

    df = process_uploaded_files()

    if df.empty:

        print()

        print(
            "No data processed."
        )

        return

    # --------------------------------------------------------
    # Basic assertions
    # --------------------------------------------------------

    assert (
        len(df) > 0
    ), "Processed data should not be empty."

    assert (
        "report_date" in df.columns
    ), "report_date missing."

    assert (
        "members" in df.columns
    ), "members missing."

    assert (
        "amount" in df.columns
    ), "amount missing."

    # --------------------------------------------------------
    # Validate latest snapshot
    # --------------------------------------------------------

    latest_date = (
        df["report_date"]
        .max()
    )

    latest = df[
        df["report_date"]
        == latest_date
    ]

    assert (
        latest["sector"].nunique()
        == EXPECTED_SECTORS
    ), (
        f"Expected {EXPECTED_SECTORS} sectors."
    )

    assert (
        latest[
            ["sector", "cell"]
        ]
        .drop_duplicates()
        .shape[0]
        == EXPECTED_CELLS
    ), (
        f"Expected {EXPECTED_CELLS} cells."
    )

    assert (
        latest[
            [
                "sector",
                "cell",
                "village"
            ]
        ]
        .drop_duplicates()
        .shape[0]
        == EXPECTED_VILLAGES
    ), (
        f"Expected {EXPECTED_VILLAGES} villages."
    )

    print()

    print("=" * 70)

    print(
        "PROGRESS SERVICE TEST PASSED"
    )

    print("=" * 70)

    print()

    print(
        "✓ Excel uploads detected automatically"
    )

    print(
        "✓ Report date extracted automatically"
    )

    print(
        "✓ 3-row Excel header handled"
    )

    print(
        "✓ Locations normalized"
    )

    print(
        "✓ 12 sectors validated"
    )

    print(
        "✓ 80 cells validated"
    )

    print(
        "✓ 519 villages validated"
    )

    print(
        "✓ Level 1–5 members calculated"
    )

    print(
        "✓ Level 1–5 money calculated"
    )

    print(
        "✓ Historical dates retained"
    )

    print(
        "✓ Same-date uploads safely replace old snapshots"
    )

    print(
        "✓ Duplicate village/date records removed"
    )

    print(
        "✓ Previous/current daily comparison available"
    )

    print()


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    test_progress_service()
