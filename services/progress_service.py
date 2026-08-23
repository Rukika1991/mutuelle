import pandas as pd
import re

from config.settings import (
    PROGRESS_DIR,
    PROCESSED_DIR
)


class ProgressService:

    # ========================================================
    # FIND ALL PROGRESS FILES
    # ========================================================

    def find_progress_files(self):

        files = sorted(
            [
                file
                for file in PROGRESS_DIR.glob("*.xlsx")
                if not file.name.startswith("~$")
            ]
        )

        if not files:

            raise FileNotFoundError(
                f"No progress Excel files found in:\n"
                f"{PROGRESS_DIR}"
            )

        return files


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
    # CLEAN ADMINISTRATIVE DATA
    # ========================================================

    def clean_administrative_data(self, df):

        df = df.copy()

        admin_columns = [
            "province",
            "district",
            "sector",
            "cell",
            "village"
        ]

        for column in admin_columns:

            if column in df.columns:

                df[column] = (
                    df[column]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

        return df


    # ========================================================
    # DETECT REPORT DATE
    # ========================================================

    def detect_report_date(self, df, file_path):

        # ----------------------------------------------------
        # 1. TRY DATE FROM EXCEL
        # ----------------------------------------------------

        possible_date_columns = [
            "report_date",
            "date",
            "reportdate",
            "reporting_date"
        ]

        for column in possible_date_columns:

            if column in df.columns:

                dates = pd.to_datetime(
                    df[column],
                    errors="coerce"
                ).dropna()

                if not dates.empty:

                    return dates.iloc[0].normalize()


        # ----------------------------------------------------
        # 2. TRY NUMERIC DATE FROM FILE NAME
        #
        # Example:
        # progress_2026-08-22.xlsx
        # ----------------------------------------------------

        filename = file_path.stem.upper()

        match = re.search(
            r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})",
            filename
        )

        if match:

            return pd.Timestamp(
                year=int(match.group(1)),
                month=int(match.group(2)),
                day=int(match.group(3))
            )


        # ----------------------------------------------------
        # 3. TRY DAY + MONTH NAME
        #
        # Examples:
        # RUBAVU_20_AUGUST.xlsx
        # RUBAVU_21_AUGUST.xlsx
        # RUBAVU_22_AUGUST.xlsx
        # ----------------------------------------------------

        month_mapping = {
            "JANUARY": 1,
            "FEBRUARY": 2,
            "MARCH": 3,
            "APRIL": 4,
            "MAY": 5,
            "JUNE": 6,
            "JULY": 7,
            "AUGUST": 8,
            "SEPTEMBER": 9,
            "OCTOBER": 10,
            "NOVEMBER": 11,
            "DECEMBER": 12,

            "JAN": 1,
            "FEB": 2,
            "MAR": 3,
            "APR": 4,
            "JUN": 6,
            "JUL": 7,
            "AUG": 8,
            "SEP": 9,
            "SEPT": 9,
            "OCT": 10,
            "NOV": 11,
            "DEC": 12
        }


        # Find day such as _20_
        day_match = re.search(
            r"(?:^|_)(\d{1,2})(?:_|$)",
            filename
        )

        detected_day = None

        if day_match:

            detected_day = int(
                day_match.group(1)
            )


        detected_month = None

        for month_name, month_number in month_mapping.items():

            if month_name in filename:

                detected_month = month_number

                break


        if (
            detected_day is not None
            and detected_month is not None
        ):

            return pd.Timestamp(
                year=2026,
                month=detected_month,
                day=detected_day
            )


        raise ValueError(
            f"Could not determine report date for:\n"
            f"{file_path.name}"
        )


    # ========================================================
    # DETECT MEMBER COLUMNS
    # ========================================================

    def detect_level_columns(self, df):

        level_mapping = {}

        for level in range(1, 6):

            possible_columns = [

                f"level{level}_members",
                f"level_{level}_members",

                f"level{level}",
                f"level_{level}",

                f"ubudehe_{level}",

                f"category_{level}",

                f"cat{level}"
            ]

            for column in possible_columns:

                if column in df.columns:

                    level_mapping[level] = column

                    break

        return level_mapping


    # ========================================================
    # DETECT AMOUNT COLUMN
    # ========================================================

    def detect_amount_column(self, df, level):

        possible_columns = [

            f"level{level}_amount",
            f"level_{level}_amount",

            f"level{level}_payment",
            f"level_{level}_payment",

            f"level{level}_contribution",
            f"level_{level}_contribution"
        ]

        for column in possible_columns:

            if column in df.columns:

                return column

        return None


    # ========================================================
    # PROCESS ONE PROGRESS FILE
    # ========================================================

    def process_file(self, file_path):

        print("\n" + "=" * 70)

        print("Processing:")

        print(file_path.name)

        print("=" * 70)


        # ----------------------------------------------------
        # LOAD EXCEL
        # ----------------------------------------------------

        df = pd.read_excel(
            file_path
        )

        print(
            f"\nRows: {len(df):,}"
        )


        # ----------------------------------------------------
        # STANDARDIZE
        # ----------------------------------------------------

        df = self.standardize_columns(df)


        # ----------------------------------------------------
        # SHOW COLUMNS
        # ----------------------------------------------------

        print("\nAVAILABLE COLUMNS:")

        for column in df.columns:

            print(f" - {column}")


        # ----------------------------------------------------
        # CLEAN ADMIN DATA
        # ----------------------------------------------------

        df = self.clean_administrative_data(df)


        # ----------------------------------------------------
        # DETECT DATE
        # ----------------------------------------------------

        report_date = self.detect_report_date(
            df,
            file_path
        )

        print(
            f"\nReport Date: "
            f"{report_date.strftime('%Y-%m-%d')}"
        )


        # ----------------------------------------------------
        # DETECT MEMBER COLUMNS
        # ----------------------------------------------------

        level_mapping = (
            self.detect_level_columns(df)
        )

        print("\nDETECTED MEMBER COLUMNS:")

        if level_mapping:

            for level, column in level_mapping.items():

                print(
                    f"Level {level}: {column}"
                )

        else:

            print(
                "⚠ No standard member columns detected."
            )


        # ----------------------------------------------------
        # CREATE STANDARD MEMBER COLUMNS
        # ----------------------------------------------------

        for level in range(1, 6):

            standard_column = (
                f"level{level}_members"
            )

            if level in level_mapping:

                source_column = (
                    level_mapping[level]
                )

                df[standard_column] = (
                    pd.to_numeric(
                        df[source_column],
                        errors="coerce"
                    )
                    .fillna(0)
                )

            else:

                df[standard_column] = 0


        # ----------------------------------------------------
        # TOTAL MEMBERS
        # ----------------------------------------------------

        member_columns = [
            f"level{level}_members"
            for level in range(1, 6)
        ]

        df["total_members"] = (
            df[member_columns]
            .sum(axis=1)
        )


        # ----------------------------------------------------
        # DETECT AMOUNT COLUMNS
        # ----------------------------------------------------

        print("\nDETECTED AMOUNT COLUMNS:")

        for level in range(1, 6):

            amount_column = (
                self.detect_amount_column(
                    df,
                    level
                )
            )

            standard_column = (
                f"level{level}_amount"
            )

            if amount_column:

                print(
                    f"Level {level}: "
                    f"{amount_column}"
                )

                df[standard_column] = (
                    pd.to_numeric(
                        df[amount_column],
                        errors="coerce"
                    )
                    .fillna(0)
                )

            else:

                df[standard_column] = 0


        # ----------------------------------------------------
        # TOTAL AMOUNT
        # ----------------------------------------------------

        amount_columns = [
            f"level{level}_amount"
            for level in range(1, 6)
        ]

        df["total_amount"] = (
            df[amount_columns]
            .sum(axis=1)
        )


        # ----------------------------------------------------
        # ADD REPORT DATE
        # ----------------------------------------------------

        df["report_date"] = report_date


        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        print("\nPROGRESS SUMMARY")

        print(
            f"Total Members: "
            f"{df['total_members'].sum():,.0f}"
        )

        print(
            f"Total Amount: "
            f"{df['total_amount'].sum():,.0f} RWF"
        )


        return df


    # ========================================================
    # PROCESS ALL PROGRESS FILES
    # ========================================================

    def process_all(self):

        files = self.find_progress_files()

        all_progress = []

        for file_path in files:

            df = self.process_file(
                file_path
            )

            all_progress.append(df)


            # Save processed daily file

            report_date = (
                df["report_date"]
                .iloc[0]
            )

            output_file = (
                PROCESSED_DIR
                /
                f"progress_"
                f"{report_date.strftime('%Y-%m-%d')}"
                f"_processed.xlsx"
            )

            df.to_excel(
                output_file,
                index=False
            )

            print(
                f"\n✓ Saved:"
            )

            print(
                output_file.name
            )


        if not all_progress:

            return pd.DataFrame()


        combined = pd.concat(
            all_progress,
            ignore_index=True
        )

        return combined