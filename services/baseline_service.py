import pandas as pd

from config.settings import (
    BASELINE_DIR,
    LEVEL_CONTRIBUTIONS
)


class BaselineService:

    # ========================================================
    # FIND BASELINE FILE
    # ========================================================

    def find_baseline_file(self):

        files = list(
            BASELINE_DIR.glob(
                "*.xlsx"
            )
        )

        if not files:

            raise FileNotFoundError(
                f"No Excel file found in:\n"
                f"{BASELINE_DIR}"
            )

        return files[0]


    # ========================================================
    # LOAD BASELINE
    # ========================================================

    def load_baseline(self):

        file_path = (
            self.find_baseline_file()
        )

        print(
            f"\nLoading baseline:\n"
            f"{file_path.name}"
        )

        df = pd.read_excel(
            file_path
        )

        print(
            f"Rows: {len(df):,}"
        )

        return df


    # ========================================================
    # STANDARDIZE COLUMN NAMES
    # ========================================================

    def standardize_columns(self, df):

        df = df.copy()

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace(
                r"[\s\-/]+",
                "_",
                regex=True
            )
            .str.replace(
                r"[^a-z0-9_]",
                "",
                regex=True
            )
        )

        return df


    # ========================================================
    # VALIDATE REQUIRED COLUMNS
    # ========================================================

    def validate_columns(self, df):

        required_columns = [

            "province",
            "district",
            "sector",
            "cell",
            "village",

            "hhsize",

            "cbhi_level"
        ]

        missing = [

            column

            for column
            in required_columns

            if column not in df.columns
        ]

        if missing:

            raise ValueError(
                "\nMissing required columns:\n"
                +
                "\n".join(
                    f"- {column}"
                    for column in missing
                )
            )

        print(
            "\n✓ Required columns found."
        )

        return df


    # ========================================================
    # CLEAN ADMINISTRATIVE DATA
    # ========================================================

    def clean_administrative_data(
        self,
        df
    ):

        df = df.copy()

        admin_columns = [

            "province",
            "district",
            "sector",
            "cell",
            "village"
        ]

        for column in admin_columns:

            df[column] = (

                df[column]

                .fillna("")

                .astype(str)

                .str.strip()

                .str.upper()

            )

        return df


    # ========================================================
    # CLEAN MEMBER COUNT
    # ========================================================

    def clean_members(self, df):

        df = df.copy()

        df["hhsize"] = pd.to_numeric(
            df["hhsize"],
            errors="coerce"
        )

        invalid = (

            df["hhsize"].isna()

            |

            (
                df["hhsize"] <= 0
            )

        )

        print(
            f"\nInvalid member records: "
            f"{invalid.sum():,}"
        )

        return df


    # ========================================================
    # CLEAN CBHI LEVEL
    # ========================================================

    def clean_cbhi_level(
        self,
        df
    ):

        df = df.copy()

        level_clean = (

            df["cbhi_level"]

            .fillna("")

            .astype(str)

            .str.lower()

            .str.extract(
                r"([1-5])",
                expand=False
            )

        )

        df["cbhi_level"] = pd.to_numeric(
            level_clean,
            errors="coerce"
        )

        return df


    # ========================================================
    # CALCULATE CONTRIBUTION
    # ========================================================

    def calculate_contribution(
        self,
        df
    ):

        df = df.copy()

        df[
            "contribution_per_member"
        ] = (

            df[
                "cbhi_level"
            ]

            .map(
                LEVEL_CONTRIBUTIONS
            )

        )


        df[
            "expected_amount"
        ] = (

            df[
                "hhsize"
            ]

            *

            df[
                "contribution_per_member"
            ]

        )

        df[
            "expected_amount"
        ] = (

            df[
                "expected_amount"
            ]

            .fillna(0)

        )

        return df


    # ========================================================
    # CREATE VILLAGE BASELINE
    # ========================================================

    def create_village_baseline(
        self,
        df
    ):

        village_baseline = (

            df

            .groupby(

                [
                    "province",
                    "district",
                    "sector",
                    "cell",
                    "village"
                ],

                dropna=False

            )

            .agg(

                baseline_members=(

                    "hhsize",

                    "sum"

                ),

                expected_amount=(

                    "expected_amount",

                    "sum"

                )

            )

            .reset_index()

        )


        # ================================================
        # MEMBERS BY CBHI LEVEL
        # ================================================

        for level in [

            1,
            2,
            3,
            4,
            5

        ]:

            level_data = (

                df[
                    df[
                        "cbhi_level"
                    ]
                    == level
                ]

                .groupby(

                    [
                        "province",
                        "district",
                        "sector",
                        "cell",
                        "village"
                    ],

                    dropna=False

                )

                .agg(

                    **{

                        f"level{level}_members":

                        (
                            "hhsize",
                            "sum"
                        ),

                        f"level{level}_expected_amount":

                        (
                            "expected_amount",
                            "sum"
                        )

                    }

                )

                .reset_index()

            )


            village_baseline = (

                village_baseline

                .merge(

                    level_data,

                    on=[

                        "province",
                        "district",
                        "sector",
                        "cell",
                        "village"

                    ],

                    how="left"

                )

            )


        # ================================================
        # REPLACE MISSING VALUES
        # ================================================

        numeric_columns = (

            village_baseline

            .select_dtypes(
                include="number"
            )

            .columns

        )

        village_baseline[
            numeric_columns
        ] = (

            village_baseline[
                numeric_columns
            ]

            .fillna(0)

        )


        return village_baseline


    # ========================================================
    # PROCESS COMPLETE BASELINE
    # ========================================================

    def process(self):

        df = self.load_baseline()

        df = self.standardize_columns(df)

        df = self.validate_columns(df)

        df = self.clean_administrative_data(df)

        df = self.clean_members(df)

        df = self.clean_cbhi_level(df)

        df = self.calculate_contribution(df)

        village_baseline = (
            self.create_village_baseline(df)
        )

        return df, village_baseline