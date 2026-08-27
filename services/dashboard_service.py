from pathlib import Path
import pandas as pd
import re


# ============================================================
# PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

BASELINE_FILE = (
    BASE_DIR
    / "data"
    / "baseline"
    / "RubavunewAugust.xlsx"
)

PROGRESS_FILE = (
    BASE_DIR
    / "output"
    / "processed_progress.xlsx"
)


# ============================================================
# EXPECTED RUBAVU STRUCTURE
# ============================================================

EXPECTED_SECTORS = 12
EXPECTED_CELLS = 80
EXPECTED_VILLAGES = 519


# ============================================================
# TEXT / COLUMN CLEANING
# ============================================================

def clean_column_name(value):
    """
    Convert a column name to a predictable format.

    Examples:
        Total Amount -> total_amount
        CBHI level   -> cbhi_level
        LEVEL 1      -> level_1
    """

    value = str(value).strip().lower()

    value = re.sub(r"[\s\-/]+", "_", value)

    value = re.sub(r"[^a-z0-9_]", "", value)

    value = re.sub(r"_+", "_", value)

    return value.strip("_")


def normalize_columns(df):
    """
    Normalize all dataframe column names.
    """

    df = df.copy()

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    return df


def clean_text(value):
    """
    Clean text values while preserving the original spelling
    for display.
    """

    if pd.isna(value):
        return ""

    value = str(value)

    value = value.replace("\u00a0", " ")

    value = value.strip()

    value = re.sub(r"\s+", " ", value)

    return value


# ============================================================
# CANONICAL LOCATION NORMALIZATION
# ============================================================

def location_key(value):
    """
    Create a comparison key for location matching.

    This does NOT change the displayed name.

    Example:

        Kanyirabigog  -> kanyirabigog
        Kanyirabigogo -> kanyirabigogo
    """

    value = clean_text(value)

    if not value:
        return ""

    value = value.lower()

    # Remove accents / unusual characters if present
    value = re.sub(r"[^a-z0-9]", "", value)

    return value


# ============================================================
# KNOWN LOCATION ALIASES
# ============================================================

# The key is the name that may appear in either dataset.
# The value is the canonical name used for matching.

CELL_ALIASES = {

    # --------------------------------------------------------
    # Kanyirabigog / Kanyirabigogo
    # --------------------------------------------------------

    "kanyirabigog": "kanyirabigogo",
    "kanyirabigogo": "kanyirabigogo",

}


VILLAGE_ALIASES = {

    # --------------------------------------------------------
    # Bonde / Bondi
    # --------------------------------------------------------

    "bonde": "bondi",
    "bondi": "bondi",

    # --------------------------------------------------------
    # Rwantobo / Rwandobo
    # --------------------------------------------------------

    "rwantobo": "rwandobo",
    "rwandobo": "rwandobo",

    # --------------------------------------------------------
    # Jenda / Jende
    # --------------------------------------------------------

    "jenda": "jende",
    "jende": "jende",

}


SECTOR_ALIASES = {}


# ============================================================
# APPLY LOCATION ALIAS
# ============================================================

def canonical_location(value, aliases=None):
    """
    Return a canonical location key.

    The original location name is NOT changed.
    """

    key = location_key(value)

    if not key:
        return ""

    if aliases and key in aliases:
        return aliases[key]

    return key


# ============================================================
# NUMERIC CLEANING
# ============================================================

def safe_numeric(series):
    """
    Safely convert values to numeric.
    """

    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)


# ============================================================
# DASHBOARD SERVICE
# ============================================================

class DashboardService:

    def __init__(
        self,
        baseline_file=BASELINE_FILE,
        progress_file=PROGRESS_FILE
    ):

        self.baseline_file = Path(baseline_file)

        self.progress_file = Path(progress_file)

        self._baseline = None

        self._progress = None


    # ========================================================
    # LOAD BASELINE
    # ========================================================

    def load_baseline(self):

        if self._baseline is not None:
            return self._baseline.copy()

        if not self.baseline_file.exists():

            raise FileNotFoundError(
                f"Baseline file not found:\n"
                f"{self.baseline_file}"
            )

        baseline = pd.read_excel(
            self.baseline_file
        )

        baseline = normalize_columns(
            baseline
        )

        self._baseline = baseline.copy()

        return baseline


    # ========================================================
    # LOAD PROGRESS
    # ========================================================

    def load_progress(self):

        if self._progress is not None:
            return self._progress.copy()

        if not self.progress_file.exists():

            raise FileNotFoundError(
                f"Processed progress file not found:\n"
                f"{self.progress_file}"
            )

        progress = pd.read_excel(
            self.progress_file
        )

        progress = normalize_columns(
            progress
        )

        self._progress = progress.copy()

        return progress


    # ========================================================
    # PREPARE BASELINE
    # ========================================================

    def prepare_baseline(self):

        baseline = self.load_baseline()

        required_columns = [
            "district",
            "sector",
            "cell",
            "village",
            "hhsize",
            "total_amount"
        ]

        missing = [
            column
            for column in required_columns
            if column not in baseline.columns
        ]

        if missing:

            raise ValueError(
                "Baseline missing required columns:\n"
                + ", ".join(missing)
                + "\n\nColumns found:\n\n"
                + ", ".join(
                    baseline.columns.tolist()
                )
            )

        # ----------------------------------------------------
        # CLEAN LOCATION COLUMNS
        # ----------------------------------------------------

        for column in [
            "district",
            "sector",
            "cell",
            "village"
        ]:

            baseline[column] = (
                baseline[column]
                .apply(clean_text)
            )

        # ----------------------------------------------------
        # CREATE CANONICAL LOCATION KEYS
        # ----------------------------------------------------

        baseline["district_key"] = (
            baseline["district"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    {}
                )
            )
        )

        baseline["sector_key"] = (
            baseline["sector"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    SECTOR_ALIASES
                )
            )
        )

        baseline["cell_key"] = (
            baseline["cell"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    CELL_ALIASES
                )
            )
        )

        baseline["village_key"] = (
            baseline["village"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    VILLAGE_ALIASES
                )
            )
        )

        # ----------------------------------------------------
        # NUMERIC VALUES
        # ----------------------------------------------------

        baseline["baseline_members"] = (
            safe_numeric(
                baseline["hhsize"]
            )
        )

        baseline["baseline_amount"] = (
            safe_numeric(
                baseline["total_amount"]
            )
        )

        return baseline


    # ========================================================
    # PREPARE PROGRESS
    # ========================================================

    def prepare_progress(self):

        progress = self.load_progress()

        required_columns = [
            "district",
            "sector",
            "cell",
            "village",
            "report_date"
        ]

        missing = [
            column
            for column in required_columns
            if column not in progress.columns
        ]

        if missing:

            raise ValueError(
                "Processed progress missing required "
                "columns:\n\n"
                + ", ".join(missing)
                + "\n\nColumns found:\n\n"
                + ", ".join(
                    progress.columns.tolist()
                )
            )

        # ----------------------------------------------------
        # CLEAN LOCATION COLUMNS
        # ----------------------------------------------------

        for column in [
            "district",
            "sector",
            "cell",
            "village"
        ]:

            progress[column] = (
                progress[column]
                .apply(clean_text)
            )

        # ----------------------------------------------------
        # CREATE CANONICAL LOCATION KEYS
        # ----------------------------------------------------

        progress["district_key"] = (
            progress["district"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    {}
                )
            )
        )

        progress["sector_key"] = (
            progress["sector"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    SECTOR_ALIASES
                )
            )
        )

        progress["cell_key"] = (
            progress["cell"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    CELL_ALIASES
                )
            )
        )

        progress["village_key"] = (
            progress["village"]
            .apply(
                lambda x:
                canonical_location(
                    x,
                    VILLAGE_ALIASES
                )
            )
        )

        # ----------------------------------------------------
        # CONVERT REPORT DATE
        # ----------------------------------------------------

        progress["report_date"] = (
            pd.to_datetime(
                progress["report_date"],
                errors="coerce"
            )
        )

        if progress["report_date"].isna().all():

            raise ValueError(
                "Could not read report_date values."
            )

        # ----------------------------------------------------
        # REQUIRED MEMBER COLUMNS
        # ----------------------------------------------------

        member_columns = [
            "level_1_members",
            "level_2_members",
            "level_3_members",
            "level_4_members",
            "level_5_members"
        ]

        # ----------------------------------------------------
        # REQUIRED MONEY COLUMNS
        # ----------------------------------------------------

        amount_columns = [
            "level_1_amount",
            "level_2_amount",
            "level_3_amount",
            "level_4_amount",
            "level_5_amount"
        ]

        missing_members = [
            column
            for column in member_columns
            if column not in progress.columns
        ]

        missing_amounts = [
            column
            for column in amount_columns
            if column not in progress.columns
        ]

        if missing_members or missing_amounts:

            missing = (
                missing_members
                +
                missing_amounts
            )

            raise ValueError(
                "Processed progress missing required "
                "member/money columns:\n\n"
                + ", ".join(missing)
                + "\n\nColumns found:\n\n"
                + ", ".join(
                    progress.columns.tolist()
                )
            )

        # ----------------------------------------------------
        # CLEAN MEMBER COLUMNS
        # ----------------------------------------------------

        for column in member_columns:

            progress[column] = (
                safe_numeric(
                    progress[column]
                )
            )

        # ----------------------------------------------------
        # CLEAN MONEY COLUMNS
        # ----------------------------------------------------

        for column in amount_columns:

            progress[column] = (
                safe_numeric(
                    progress[column]
                )
            )

        # ----------------------------------------------------
        # CALCULATE TOTAL MEMBERS
        # ----------------------------------------------------

        progress["total_members"] = (
            progress[member_columns]
            .sum(axis=1)
        )

        # ----------------------------------------------------
        # CALCULATE TOTAL MONEY
        # ----------------------------------------------------

        progress["total_amount"] = (
            progress[amount_columns]
            .sum(axis=1)
        )

        return progress


    # ========================================================
    # VALIDATE RUBAVU STRUCTURE
    # ========================================================

    def validate_structure(self):

        baseline = self.prepare_baseline()

        progress = self.prepare_progress()

        # ----------------------------------------------------
        # BASELINE STRUCTURE
        # ----------------------------------------------------

        baseline_sectors = (
            baseline["sector_key"]
            .nunique()
        )

        baseline_cells = (
            baseline[
                [
                    "sector_key",
                    "cell_key"
                ]
            ]
            .drop_duplicates()
            .shape[0]
        )

        baseline_villages = (
            baseline[
                [
                    "sector_key",
                    "cell_key",
                    "village_key"
                ]
            ]
            .drop_duplicates()
            .shape[0]
        )

        # ----------------------------------------------------
        # PROGRESS STRUCTURE
        # ----------------------------------------------------

        progress_sectors = (
            progress["sector_key"]
            .nunique()
        )

        progress_cells = (
            progress[
                [
                    "sector_key",
                    "cell_key"
                ]
            ]
            .drop_duplicates()
            .shape[0]
        )

        progress_villages = (
            progress[
                [
                    "sector_key",
                    "cell_key",
                    "village_key"
                ]
            ]
            .drop_duplicates()
            .shape[0]
        )

        return {

            "baseline": {
                "sectors": baseline_sectors,
                "cells": baseline_cells,
                "villages": baseline_villages
            },

            "progress": {
                "sectors": progress_sectors,
                "cells": progress_cells,
                "villages": progress_villages
            }

        }


    # ========================================================
    # GET LATEST TWO DATES
    # ========================================================

    def get_latest_dates(self):

        progress = self.prepare_progress()

        dates = (
            progress["report_date"]
            .dropna()
            .drop_duplicates()
            .sort_values()
            .tolist()
        )

        if len(dates) < 2:

            raise ValueError(
                "At least two report dates are required "
                "to calculate today's collection."
            )

        previous_date = dates[-2]

        current_date = dates[-1]

        return previous_date, current_date


    # ========================================================
    # CALCULATE INDICATORS
    # ========================================================

    @staticmethod
    def calculate_indicators(
        baseline_members,
        previous_members,
        current_members,
        baseline_amount,
        previous_amount,
        current_amount
    ):

        # ----------------------------------------------------
        # MEMBERS
        # ----------------------------------------------------

        today_members = (
            current_members
            -
            previous_members
        )

        remaining_members = (
            baseline_members
            -
            current_members
        )

        member_progress = (
            (
                current_members
                /
                baseline_members
                *
                100
            )
            if baseline_members > 0
            else 0
        )

        # ----------------------------------------------------
        # MONEY
        # ----------------------------------------------------

        today_amount = (
            current_amount
            -
            previous_amount
        )

        remaining_amount = (
            baseline_amount
            -
            current_amount
        )

        money_progress = (
            (
                current_amount
                /
                baseline_amount
                *
                100
            )
            if baseline_amount > 0
            else 0
        )

        return {

            # MEMBERS

            "baseline_members":
                baseline_members,

            "previous_members":
                previous_members,

            "current_members":
                current_members,

            "today_members":
                today_members,

            "remaining_members":
                remaining_members,

            "member_progress_percent":
                member_progress,

            # MONEY

            "baseline_amount":
                baseline_amount,

            "previous_amount":
                previous_amount,

            "current_amount":
                current_amount,

            "today_amount":
                today_amount,

            "remaining_amount":
                remaining_amount,

            "money_progress_percent":
                money_progress
        }


    # ========================================================
    # DISTRICT DASHBOARD
    # ========================================================

    def get_district_dashboard(self):

        baseline = self.prepare_baseline()

        progress = self.prepare_progress()

        previous_date, current_date = (
            self.get_latest_dates()
        )

        # ----------------------------------------------------
        # BASELINE TOTALS
        # ----------------------------------------------------

        baseline_members = (
            baseline["baseline_members"]
            .sum()
        )

        baseline_amount = (
            baseline["baseline_amount"]
            .sum()
        )

        # ----------------------------------------------------
        # PREVIOUS PROGRESS
        # ----------------------------------------------------

        previous = (
            progress[
                progress["report_date"]
                ==
                previous_date
            ]
        )

        previous_members = (
            previous["total_members"]
            .sum()
        )

        previous_amount = (
            previous["total_amount"]
            .sum()
        )

        # ----------------------------------------------------
        # CURRENT PROGRESS
        # ----------------------------------------------------

        current = (
            progress[
                progress["report_date"]
                ==
                current_date
            ]
        )

        current_members = (
            current["total_members"]
            .sum()
        )

        current_amount = (
            current["total_amount"]
            .sum()
        )

        # ----------------------------------------------------
        # CALCULATE
        # ----------------------------------------------------

        indicators = (
            self.calculate_indicators(

                baseline_members=
                    baseline_members,

                previous_members=
                    previous_members,

                current_members=
                    current_members,

                baseline_amount=
                    baseline_amount,

                previous_amount=
                    previous_amount,

                current_amount=
                    current_amount
            )
        )

        indicators["previous_date"] = previous_date

        indicators["current_date"] = current_date

        return indicators


    # ========================================================
    # GENERIC PERFORMANCE BY LOCATION LEVEL
    # ========================================================

    def get_performance_by_level(self, level):

        allowed_levels = [
            "sector",
            "cell",
            "village"
        ]

        if level not in allowed_levels:

            raise ValueError(
                f"Invalid level: {level}. "
                f"Use one of: "
                f"{', '.join(allowed_levels)}"
            )

        baseline = self.prepare_baseline()

        progress = self.prepare_progress()

        previous_date, current_date = (
            self.get_latest_dates()
        )

        # ----------------------------------------------------
        # GROUPING KEYS
        # ----------------------------------------------------

        if level == "sector":

            group_columns = [
                "sector_key"
            ]

            display_columns = [
                "sector"
            ]

        elif level == "cell":

            group_columns = [
                "sector_key",
                "cell_key"
            ]

            display_columns = [
                "sector",
                "cell"
            ]

        else:

            group_columns = [
                "sector_key",
                "cell_key",
                "village_key"
            ]

            display_columns = [
                "sector",
                "cell",
                "village"
            ]

        # ----------------------------------------------------
        # BASELINE
        # ----------------------------------------------------

        baseline_summary = (
            baseline
            .groupby(
                group_columns,
                as_index=False,
                dropna=False
            )
            .agg(
                baseline_members=(
                    "baseline_members",
                    "sum"
                ),

                baseline_amount=(
                    "baseline_amount",
                    "sum"
                )
            )
        )

        # ----------------------------------------------------
        # PREVIOUS
        # ----------------------------------------------------

        previous = (
            progress[
                progress["report_date"]
                ==
                previous_date
            ]
            .groupby(
                group_columns,
                as_index=False,
                dropna=False
            )
            .agg(
                previous_members=(
                    "total_members",
                    "sum"
                ),

                previous_amount=(
                    "total_amount",
                    "sum"
                )
            )
        )

        # ----------------------------------------------------
        # CURRENT
        # ----------------------------------------------------

        current = (
            progress[
                progress["report_date"]
                ==
                current_date
            ]
            .groupby(
                group_columns,
                as_index=False,
                dropna=False
            )
            .agg(
                current_members=(
                    "total_members",
                    "sum"
                ),

                current_amount=(
                    "total_amount",
                    "sum"
                )
            )
        )

        # ----------------------------------------------------
        # DISPLAY NAMES FROM BASELINE
        # ----------------------------------------------------

        display_summary = (
            baseline
            .groupby(
                group_columns,
                as_index=False,
                dropna=False
            )
            .agg(
                **{
                    column: (
                        column,
                        "first"
                    )
                    for column in display_columns
                }
            )
        )

        # ----------------------------------------------------
        # MERGE
        # ----------------------------------------------------

        result = (
            display_summary
            .merge(
                baseline_summary,
                on=group_columns,
                how="left"
            )
            .merge(
                previous,
                on=group_columns,
                how="left"
            )
            .merge(
                current,
                on=group_columns,
                how="left"
            )
        )

        # ----------------------------------------------------
        # FILL EMPTY VALUES
        # ----------------------------------------------------

        numeric_columns = [
            "baseline_members",
            "baseline_amount",
            "previous_members",
            "previous_amount",
            "current_members",
            "current_amount"
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

        # ----------------------------------------------------
        # MEMBERS
        # ----------------------------------------------------

        result["today_members"] = (
            result["current_members"]
            -
            result["previous_members"]
        )

        result["remaining_members"] = (
            result["baseline_members"]
            -
            result["current_members"]
        )

        result["member_progress_percent"] = (
            (
                result["current_members"]
                /
                result["baseline_members"]
                .replace(0, pd.NA)
                *
                100
            )
            .fillna(0)
        )

        # ----------------------------------------------------
        # MONEY
        # ----------------------------------------------------

        result["today_amount"] = (
            result["current_amount"]
            -
            result["previous_amount"]
        )

        result["remaining_amount"] = (
            result["baseline_amount"]
            -
            result["current_amount"]
        )

        result["money_progress_percent"] = (
            (
                result["current_amount"]
                /
                result["baseline_amount"]
                .replace(0, pd.NA)
                *
                100
            )
            .fillna(0)
        )

        # ----------------------------------------------------
        # ADD RANK
        # ----------------------------------------------------

        result = (
            result
            .sort_values(
                [
                    "money_progress_percent",
                    "member_progress_percent"
                ],
                ascending=False
            )
            .reset_index(drop=True)
        )

        result.insert(
            0,
            "Rank",
            range(1, len(result) + 1)
        )

        # ----------------------------------------------------
        # FINAL COLUMN ORDER
        # ----------------------------------------------------

        if level == "sector":

            result = result[
                [
                    "Rank",
                    "sector",

                    "baseline_members",
                    "baseline_amount",

                    "current_members",
                    "current_amount",

                    "previous_members",
                    "previous_amount",

                    "today_members",
                    "today_amount",

                    "remaining_members",
                    "remaining_amount",

                    "member_progress_percent",
                    "money_progress_percent"
                ]
            ]

        elif level == "cell":

            result = result[
                [
                    "Rank",
                    "sector",
                    "cell",

                    "baseline_members",
                    "baseline_amount",

                    "current_members",
                    "current_amount",

                    "previous_members",
                    "previous_amount",

                    "today_members",
                    "today_amount",

                    "remaining_members",
                    "remaining_amount",

                    "member_progress_percent",
                    "money_progress_percent"
                ]
            ]

        else:

            result = result[
                [
                    "Rank",
                    "sector",
                    "cell",
                    "village",

                    "baseline_members",
                    "baseline_amount",

                    "current_members",
                    "current_amount",

                    "previous_members",
                    "previous_amount",

                    "today_members",
                    "today_amount",

                    "remaining_members",
                    "remaining_amount",

                    "member_progress_percent",
                    "money_progress_percent"
                ]
            ]

        return result


    # ========================================================
    # SECTOR PERFORMANCE
    # ========================================================

    def get_sector_performance(self):

        return self.get_performance_by_level(
            "sector"
        )


    # ========================================================
    # CELL PERFORMANCE
    # ========================================================

    def get_cell_performance(self):

        return self.get_performance_by_level(
            "cell"
        )


    # ========================================================
    # VILLAGE PERFORMANCE
    # ========================================================

    def get_village_performance(self):

        return self.get_performance_by_level(
            "village"
        )


# ============================================================
# TEST
# ============================================================

def test_dashboard_service():

    print()

    print("=" * 70)

    print(
        "RUBAVU DASHBOARD SERVICE TEST"
    )

    print("=" * 70)

    print()

    print(
        "Loading baseline and processed progress..."
    )

    service = DashboardService()

    # --------------------------------------------------------
    # VALIDATE STRUCTURE
    # --------------------------------------------------------

    validation = (
        service.validate_structure()
    )

    print()

    print("DATA VALIDATION")

    print()

    print("-" * 70)

    baseline = (
        service.prepare_baseline()
    )

    progress = (
        service.prepare_progress()
    )

    print()

    print(
        f"Baseline rows: "
        f"{len(baseline):,}"
    )

    print(
        f"Progress rows: "
        f"{len(progress):,}"
    )

    print()

    print("BASELINE STRUCTURE")

    print()

    print(
        f"Sectors: "
        f"{validation['baseline']['sectors']}"
    )

    print(
        f"Cells: "
        f"{validation['baseline']['cells']}"
    )

    print(
        f"Villages: "
        f"{validation['baseline']['villages']}"
    )

    print()

    print("PROGRESS STRUCTURE")

    print()

    print(
        f"Sectors: "
        f"{validation['progress']['sectors']}"
    )

    print(
        f"Cells: "
        f"{validation['progress']['cells']}"
    )

    print(
        f"Villages: "
        f"{validation['progress']['villages']}"
    )

    # --------------------------------------------------------
    # DISTRICT
    # --------------------------------------------------------

    dashboard = (
        service.get_district_dashboard()
    )

    print()

    print("=" * 70)

    print(
        "DISTRICT DASHBOARD"
    )

    print("=" * 70)

    print()

    print(
        f"Previous date: "
        f"{dashboard['previous_date'].date()}"
    )

    print(
        f"Current date:  "
        f"{dashboard['current_date'].date()}"
    )

    # --------------------------------------------------------
    # MEMBERS
    # --------------------------------------------------------

    print()

    print("-" * 70)

    print(
        "MEMBER PERFORMANCE"
    )

    print("-" * 70)

    print()

    print(
        f"Baseline members: "
        f"{dashboard['baseline_members']:,.0f}"
    )

    print(
        f"Previous members: "
        f"{dashboard['previous_members']:,.0f}"
    )

    print(
        f"Current members:  "
        f"{dashboard['current_members']:,.0f}"
    )

    print(
        f"Today members:    "
        f"{dashboard['today_members']:,.0f}"
    )

    print(
        f"Remaining members:"
        f" {dashboard['remaining_members']:,.0f}"
    )

    print(
        f"Member progress:  "
        f"{dashboard['member_progress_percent']:.2f}%"
    )

    # --------------------------------------------------------
    # MONEY
    # --------------------------------------------------------

    print()

    print("-" * 70)

    print(
        "MONEY PERFORMANCE"
    )

    print("-" * 70)

    print()

    print(
        f"Baseline amount: "
        f"{dashboard['baseline_amount']:,.0f}"
    )

    print(
        f"Previous paid:   "
        f"{dashboard['previous_amount']:,.0f}"
    )

    print(
        f"Current paid:    "
        f"{dashboard['current_amount']:,.0f}"
    )

    print(
        f"Today collected: "
        f"{dashboard['today_amount']:,.0f}"
    )

    print(
        f"Remaining:       "
        f"{dashboard['remaining_amount']:,.0f}"
    )

    print(
        f"Money progress:  "
        f"{dashboard['money_progress_percent']:.2f}%"
    )

    # --------------------------------------------------------
    # SECTOR PERFORMANCE
    # --------------------------------------------------------

    print()

    print("=" * 70)

    print(
        "SECTOR PERFORMANCE"
    )

    print("=" * 70)

    print()

    sectors = (
        service.get_sector_performance()
    )

    print(
        sectors.to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # CELL PERFORMANCE
    # --------------------------------------------------------

    print()

    print("=" * 70)

    print(
        "CELL PERFORMANCE SAMPLE"
    )

    print("=" * 70)

    print()

    cells = (
        service.get_cell_performance()
    )

    print(
        cells.head(10).to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # VILLAGE PERFORMANCE
    # --------------------------------------------------------

    print()

    print("=" * 70)

    print(
        "VILLAGE PERFORMANCE SAMPLE"
    )

    print("=" * 70)

    print()

    villages = (
        service.get_village_performance()
    )

    print(
        villages.head(10).to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # CHECK KNOWN MATCHES
    # --------------------------------------------------------

    print()

    print("=" * 70)

    print(
        "LOCATION MATCHING CHECK"
    )

    print("=" * 70)

    print()

    print(
        "Known aliases:"
    )

    print(
        "  Cell:    Kanyirabigog <-> Kanyirabigogo"
    )

    print(
        "  Village: Bonde <-> Bondi"
    )

    print(
        "  Village: Rwantobo <-> Rwandobo"
    )

    print(
        "  Village: Jenda <-> Jende"
    )

    print()

    # --------------------------------------------------------
    # TEST ALIAS KEYS
    # --------------------------------------------------------

    test_locations = [
        ("Kanyirabigog", CELL_ALIASES),
        ("Kanyirabigogo", CELL_ALIASES),
        ("Bonde", VILLAGE_ALIASES),
        ("Bondi", VILLAGE_ALIASES),
        ("Rwantobo", VILLAGE_ALIASES),
        ("Rwandobo", VILLAGE_ALIASES),
        ("Jenda", VILLAGE_ALIASES),
        ("Jende", VILLAGE_ALIASES),
    ]

    for name, aliases in test_locations:

        print(
            f"{name:20s} -> "
            f"{canonical_location(name, aliases)}"
        )

    print()

    print("=" * 70)

    print(
        "DASHBOARD SERVICE TEST PASSED"
    )

    print("=" * 70)


# ============================================================
# RUN TEST
# ============================================================

if __name__ == "__main__":

    test_dashboard_service()