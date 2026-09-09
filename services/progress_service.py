
# ============================================================
# AUTOMATIC PROCESSOR
# ============================================================

def process_uploaded_files():
    """
    Main automatic processing function.

    Every Excel file in data/uploads is inspected.

    Invalid Excel files are SKIPPED instead of stopping the
    entire processing operation.

    Historical dates are retained.

    Same-date snapshots replace the old snapshot.

    This means:
        - One bad Excel does NOT crash the dashboard
        - Valid Excel files continue processing
        - Invalid files are reported in the log
        - Dashboard continues normally
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

    # Keep track of files that could not be processed
    skipped_files = []

    for file_path in files:

        print()

        print("-" * 70)

        print(
            f"Processing file: {file_path.name}"
        )

        try:

            # ------------------------------------------------
            # Extract report date
            # ------------------------------------------------

            report_date = extract_report_date(
                file_path
            )

            print(
                f"Report date detected: "
                f"{report_date.strftime('%Y-%m-%d')}"
            )

            # ------------------------------------------------
            # IMPORTANT:
            #
            # We still read the file even if the date already
            # exists.
            #
            # If a corrected report is uploaded for the same
            # date, it replaces the previous snapshot.
            # ------------------------------------------------

            df = process_file(
                file_path
            )

            new_data_frames.append(
                df
            )

            print(
                f"✓ Successfully processed: "
                f"{file_path.name}"
            )

        except Exception as exc:

            # ------------------------------------------------
            # IMPORTANT:
            #
            # DO NOT raise the error.
            #
            # Skip this file and continue with the next Excel.
            # ------------------------------------------------

            print()

            print(
                f"ERROR processing "
                f"{file_path.name}"
            )

            print(
                f"→ {str(exc)}"
            )

            print(
                "→ SKIP this file"
            )

            print(
                "→ Continue processing remaining Excel files"
            )

            skipped_files.append(
                {
                    "file": file_path.name,
                    "error": str(exc),
                }
            )

            # Move to the NEXT file
            continue

    # --------------------------------------------------------
    # Processing summary
    # --------------------------------------------------------

    print()

    print("=" * 70)

    print(
        "FILE PROCESSING SUMMARY"
    )

    print("=" * 70)

    print()

    print(
        f"Total Excel files found: {len(files)}"
    )

    print(
        f"Successfully processed: "
        f"{len(new_data_frames)}"
    )

    print(
        f"Skipped files: "
        f"{len(skipped_files)}"
    )

    # --------------------------------------------------------
    # Show skipped files
    # --------------------------------------------------------

    if skipped_files:

        print()

        print(
            "Skipped Excel files:"
        )

        for item in skipped_files:

            print(
                f"  - {item['file']}"
            )

            print(
                f"    Reason: {item['error']}"
            )

    print()

    # --------------------------------------------------------
    # Combine all valid uploads
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

