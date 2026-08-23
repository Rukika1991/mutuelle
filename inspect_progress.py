import pandas as pd
from pathlib import Path

# ============================================================
# CONFIGURATION
# ============================================================

FILE_PATH = "data/progress/RUBAVU_22_AUGUST.xlsx"   # CHANGE THIS TO YOUR FILE NAME
OUTPUT_FILE = "Rubavu_progress_cleaned.xlsx"


# ============================================================
# COLUMN STRUCTURE
# ============================================================

COLUMNS = [
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
    "level5_amount",
]


NUMERIC_COLUMNS = [
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
    "level5_amount",
]


# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel(file_path):

    print("\n" + "=" * 80)
    print("LOADING PROGRESS EXCEL FILE")
    print("=" * 80)

    print(f"\nFile: {file_path}")

    df_raw = pd.read_excel(
        file_path,
        header=None
    )

    print(f"Raw shape: {df_raw.shape}")

    return df_raw


# ============================================================
# CLEAN DATA
# ============================================================

def clean_dataframe(df_raw):

    print("\n" + "=" * 80)
    print("CLEANING DATA")
    print("=" * 80)

    # --------------------------------------------------------
    # Actual village data starts on row 3
    # --------------------------------------------------------

    df = df_raw.iloc[3:].copy()

    # Keep only expected 19 columns
    df = df.iloc[:, :19]

    # Rename columns
    df.columns = COLUMNS

    # --------------------------------------------------------
    # Clean text columns
    # --------------------------------------------------------

    text_columns = [
        "district",
        "sector",
        "cell",
        "village"
    ]

    for col in text_columns:

        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        # Convert strings like "NAN" to empty
        df[col] = df[col].replace(
            ["NAN", "NONE", ""],
            pd.NA
        )

    # --------------------------------------------------------
    # Remove rows without village
    # --------------------------------------------------------

    before = len(df)

    df = df[
        df["village"].notna()
    ].copy()

    removed = before - len(df)

    print(f"Rows removed: {removed}")

    # --------------------------------------------------------
    # Convert numeric columns
    # --------------------------------------------------------

    for col in NUMERIC_COLUMNS:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)

    # --------------------------------------------------------
    # Ensure integer fields are integers
    # --------------------------------------------------------

    household_columns = [
        "level1_households",
        "level2_households",
        "level3_households",
        "level4_households",
        "level5_households",
    ]

    member_columns = [
        "level1_members",
        "level2_members",
        "level3_members",
        "level4_members",
        "level5_members",
    ]

    for col in household_columns + member_columns:

        df[col] = df[col].astype(int)

    # --------------------------------------------------------
    # Calculate totals
    # --------------------------------------------------------

    df["total_households"] = (
        df["level1_households"]
        + df["level2_households"]
        + df["level3_households"]
        + df["level4_households"]
        + df["level5_households"]
    )

    df["total_members"] = (
        df["level1_members"]
        + df["level2_members"]
        + df["level3_members"]
        + df["level4_members"]
        + df["level5_members"]
    )

    df["total_amount"] = (
        df["level1_amount"]
        + df["level2_amount"]
        + df["level3_amount"]
        + df["level4_amount"]
        + df["level5_amount"]
    )

    # --------------------------------------------------------
    # Calculate average amount per household
    # --------------------------------------------------------

    df["amount_per_household"] = 0.0

    mask = df["total_households"] > 0

    df.loc[mask, "amount_per_household"] = (
        df.loc[mask, "total_amount"]
        / df.loc[mask, "total_households"]
    )

    # --------------------------------------------------------
    # Sort geographically
    # --------------------------------------------------------

    df = df.sort_values(
        [
            "district",
            "sector",
            "cell",
            "village"
        ]
    ).reset_index(drop=True)

    print(f"Clean dataframe shape: {df.shape}")

    return df


# ============================================================
# SUMMARY FUNCTION
# ============================================================

def create_summary(df, group_columns):

    summary = (
        df.groupby(group_columns, dropna=False)
        .agg(
            villages=("village", "nunique"),

            total_households=("total_households", "sum"),
            total_members=("total_members", "sum"),
            total_amount=("total_amount", "sum"),

            level1_households=("level1_households", "sum"),
            level1_members=("level1_members", "sum"),
            level1_amount=("level1_amount", "sum"),

            level2_households=("level2_households", "sum"),
            level2_members=("level2_members", "sum"),
            level2_amount=("level2_amount", "sum"),

            level3_households=("level3_households", "sum"),
            level3_members=("level3_members", "sum"),
            level3_amount=("level3_amount", "sum"),

            level4_households=("level4_households", "sum"),
            level4_members=("level4_members", "sum"),
            level4_amount=("level4_amount", "sum"),

            level5_households=("level5_households", "sum"),
            level5_members=("level5_members", "sum"),
            level5_amount=("level5_amount", "sum"),
        )
        .reset_index()
    )

    # --------------------------------------------------------
    # Amount per household
    # --------------------------------------------------------

    summary["amount_per_household"] = 0.0

    mask = summary["total_households"] > 0

    summary.loc[mask, "amount_per_household"] = (
        summary.loc[mask, "total_amount"]
        / summary.loc[mask, "total_households"]
    )

    return summary


# ============================================================
# PRINT SUMMARY
# ============================================================

def print_summary(summary, title, group_columns):

    print("\n" + "=" * 80)
    print(title)
    print("=" * 80)

    display_columns = (
        group_columns
        + [
            "villages",
            "total_households",
            "total_members",
            "total_amount",
        ]
    )

    print(
        summary[display_columns].to_string(
            index=False
        )
    )


# ============================================================
# EXPORT TO EXCEL
# ============================================================

def export_results(
    clean_df,
    district_summary,
    sector_summary,
    cell_summary,
    village_summary
):

    print("\n" + "=" * 80)
    print("EXPORTING RESULTS")
    print("=" * 80)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # Clean village-level data
        # ----------------------------------------------------

        clean_df.to_excel(
            writer,
            sheet_name="Clean Data",
            index=False
        )

        # ----------------------------------------------------
        # District
        # ----------------------------------------------------

        district_summary.to_excel(
            writer,
            sheet_name="District Summary",
            index=False
        )

        # ----------------------------------------------------
        # Sector
        # ----------------------------------------------------

        sector_summary.to_excel(
            writer,
            sheet_name="Sector Summary",
            index=False
        )

        # ----------------------------------------------------
        # Cell
        # ----------------------------------------------------

        cell_summary.to_excel(
            writer,
            sheet_name="Cell Summary",
            index=False
        )

        # ----------------------------------------------------
        # Village
        # ----------------------------------------------------

        village_summary.to_excel(
            writer,
            sheet_name="Village Summary",
            index=False
        )

    print(f"\nOutput file created:")
    print(OUTPUT_FILE)


# ============================================================
# MAIN
# ============================================================

def main():

    # --------------------------------------------------------
    # Check file
    # --------------------------------------------------------

    if not Path(FILE_PATH).exists():

        print("\nERROR")
        print("-" * 80)
        print(f"File not found: {FILE_PATH}")
        print("\nAvailable Excel files:")

        for file in Path(".").glob("*.xlsx"):
            print(f"  - {file}")

        return

    # --------------------------------------------------------
    # Load
    # --------------------------------------------------------

    raw_df = load_excel(FILE_PATH)

    # --------------------------------------------------------
    # Clean
    # --------------------------------------------------------

    df = clean_dataframe(raw_df)

    # --------------------------------------------------------
    # Create summaries
    # --------------------------------------------------------

    district_summary = create_summary(
        df,
        ["district"]
    )

    sector_summary = create_summary(
        df,
        ["district", "sector"]
    )

    cell_summary = create_summary(
        df,
        ["district", "sector", "cell"]
    )

    village_summary = create_summary(
        df,
        ["district", "sector", "cell", "village"]
    )

    # --------------------------------------------------------
    # Print District
    # --------------------------------------------------------

    print_summary(
        district_summary,
        "DISTRICT SUMMARY",
        ["district"]
    )

    # --------------------------------------------------------
    # Print Sector
    # --------------------------------------------------------

    print_summary(
        sector_summary,
        "SECTOR SUMMARY",
        ["district", "sector"]
    )

    # --------------------------------------------------------
    # Print Cell
    # --------------------------------------------------------

    print_summary(
        cell_summary,
        "CELL SUMMARY",
        ["district", "sector", "cell"]
    )

    # --------------------------------------------------------
    # Print first 20 villages
    # --------------------------------------------------------

    print("\n" + "=" * 80)
    print("VILLAGE SUMMARY - FIRST 20")
    print("=" * 80)

    print(
        village_summary[
            [
                "district",
                "sector",
                "cell",
                "village",
                "total_households",
                "total_members",
                "total_amount",
            ]
        ]
        .head(20)
        .to_string(index=False)
    )

    # --------------------------------------------------------
    # Overall totals
    # --------------------------------------------------------

    print("\n" + "=" * 80)
    print("OVERALL RUBAVU TOTAL")
    print("=" * 80)

    total_households = df["total_households"].sum()
    total_members = df["total_members"].sum()
    total_amount = df["total_amount"].sum()

    print(f"Villages       : {df['village'].nunique():,}")
    print(f"Cells          : {df['cell'].nunique():,}")
    print(f"Sectors        : {df['sector'].nunique():,}")
    print(f"Households     : {total_households:,}")
    print(f"Members        : {total_members:,}")
    print(f"Amount         : {total_amount:,.0f} RWF")

    if total_households > 0:

        print(
            f"Amount/HH      : "
            f"{total_amount / total_households:,.2f} RWF"
        )

    # --------------------------------------------------------
    # Export
    # --------------------------------------------------------

    export_results(
        df,
        district_summary,
        sector_summary,
        cell_summary,
        village_summary
    )

    print("\n" + "=" * 80)
    print("PROCESS COMPLETED SUCCESSFULLY")
    print("=" * 80)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()