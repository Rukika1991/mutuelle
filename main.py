import streamlit as st
import pandas as pd
from pathlib import Path
import shutil
import subprocess
import sys
import time

from services.dashboard_service import DashboardService


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Rubavu MUSA Monitoring",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"

UPLOAD_DIR = DATA_DIR / "uploads"

PROGRESS_DIR = DATA_DIR / "progress"

OUTPUT_DIR = BASE_DIR / "output"

PROCESSED_FILE = OUTPUT_DIR / "processed_progress.xlsx"


UPLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

PROGRESS_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True,
)


# ============================================================
# EXPECTED RUBAVU STRUCTURE
# ============================================================

EXPECTED_SECTORS = 12

EXPECTED_CELLS = 80

EXPECTED_VILLAGES = 519


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 0px;
    }

    .subtitle {
        font-size: 15px;
        color: #666;
        margin-bottom: 25px;
    }

    .kpi-card {
        background: white;
        border: 1px solid #dddddd;
        border-radius: 12px;
        padding: 18px;
        min-height: 125px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }

    .kpi-label {
        font-size: 14px;
        color: #666;
        font-weight: 600;
        margin-bottom: 8px;
    }

    .kpi-value {
        font-size: 27px;
        font-weight: 800;
        margin-top: 8px;
        color: #000000 !important;
    }

    .section-title {
        font-size: 23px;
        font-weight: 800;
        margin-top: 25px;
        margin-bottom: 12px;
    }

    .date-box {
        background: #f5f5f5;
        border-radius: 10px;
        padding: 12px 16px;
        margin-bottom: 20px;
    }

    .upload-box {
        border: 2px dashed #999;
        border-radius: 12px;
        padding: 20px;
        background: #fafafa;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FORMAT HELPERS
# ============================================================

def fmt_number(value):

    if pd.isna(value):
        return "0"

    try:
        return f"{float(value):,.0f}"

    except Exception:
        return "0"


def fmt_money(value):

    if pd.isna(value):
        return "0 RWF"

    try:
        return f"{float(value):,.0f} RWF"

    except Exception:
        return "0 RWF"


def fmt_percent(value):

    if pd.isna(value):
        return "0.00%"

    try:
        return f"{float(value):.2f}%"

    except Exception:
        return "0.00%"


# ============================================================
# RUN PROGRESS SERVICE
# ============================================================

def run_progress_service():

    """
    Run the existing progress_service.py.
    """

    result = subprocess.run(

        [
            sys.executable,
            "-m",
            "services.progress_service",
        ],

        cwd=str(BASE_DIR),

        capture_output=True,

        text=True,

    )

    return result


# ============================================================
# LOAD DASHBOARD
# ============================================================

def load_dashboard():

    """
    Create a fresh DashboardService.
    """

    service = DashboardService()

    district = service.get_district_dashboard()

    sectors = service.get_sector_performance()

    cells = service.get_cell_performance()

    villages = service.get_village_performance()

    return (
        service,
        district,
        sectors,
        cells,
        villages,
    )


# ============================================================
# SAFE RANK FUNCTION
# ============================================================

def add_rank(df):

    """
    Safely add ranking.

    Removes any existing Rank column first.
    """

    df = df.copy()

    rank_columns = [

        column

        for column in df.columns

        if str(column).strip().lower() == "rank"

    ]

    if rank_columns:

        df = df.drop(
            columns=rank_columns
        )

    df = df.reset_index(
        drop=True
    )

    df.insert(

        0,

        "Rank",

        range(
            1,
            len(df) + 1,
        ),

    )

    return df


# ============================================================
# FORMAT TABLE
# ============================================================

def prepare_table(
    df,
    level,
):

    """
    Prepare output tables.

    All calculations remain inside DashboardService.
    """

    table = df.copy()

    # --------------------------------------------------------
    # SAFE RANK
    # --------------------------------------------------------

    table = add_rank(
        table
    )


    # --------------------------------------------------------
    # COLUMN RENAMING
    # --------------------------------------------------------

    rename_map = {

        "sector": "Sector",

        "cell": "Cell",

        "village": "Village",

        "baseline_members": "Baseline Members",

        "baseline_amount": "Baseline Money",

        "previous_members": "Previous Members",

        "previous_amount": "Previous Paid",

        "current_members": "Current Members",

        "current_amount": "Current Paid",

        "today_members": "Today Enrolled",

        "remaining_members": "Remaining Members",

        "member_progress_percent": "Member Progress %",

        "today_amount": "Today Collected",

        "remaining_amount": "Remaining Money",

        "money_progress_percent": "Money Progress %",

    }

    table = table.rename(
        columns=rename_map
    )


    # --------------------------------------------------------
    # COLUMNS TO DISPLAY
    # --------------------------------------------------------

    if level == "sector":

        columns = [

            "Rank",

            "Sector",

            "Baseline Money",

            "Baseline Members",

            "Previous Paid",

            "Previous Members",

            "Current Paid",

            "Current Members",

            "Today Collected",

            "Today Enrolled",

            "Remaining Money",

            "Remaining Members",

            "Money Progress %",

            "Member Progress %",

        ]


    elif level == "cell":

        columns = [

            "Rank",

            "Sector",

            "Cell",

            "Baseline Money",

            "Baseline Members",

            "Previous Paid",

            "Previous Members",

            "Current Paid",

            "Current Members",

            "Today Collected",

            "Today Enrolled",

            "Remaining Money",

            "Remaining Members",

            "Money Progress %",

            "Member Progress %",

        ]


    else:

        columns = [

            "Rank",

            "Sector",

            "Cell",

            "Village",

            "Baseline Money",

            "Baseline Members",

            "Previous Paid",

            "Previous Members",

            "Current Paid",

            "Current Members",

            "Today Collected",

            "Today Enrolled",

            "Remaining Money",

            "Remaining Members",

            "Money Progress %",

            "Member Progress %",

        ]


    # --------------------------------------------------------
    # KEEP ONLY EXISTING COLUMNS
    # --------------------------------------------------------

    columns_available = [

        column

        for column in columns

        if column in table.columns

    ]

    table = table[
        columns_available
    ]

    return table


# ============================================================
# HEADER
# ============================================================

st.markdown(

    '<div class="main-title">'
    'RUBAVU MUSA MONITORING'
    '</div>',

    unsafe_allow_html=True,

)

st.markdown(

    '<div class="subtitle">'
    'Daily Mutuelle de Santé collection and enrolment monitoring'
    '</div>',

    unsafe_allow_html=True,

)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header(
        "📂 Upload Daily Report"
    )

    st.write(
        "Upload the latest Rubavu Mutuelle Excel report."
    )

    uploaded_file = st.file_uploader(

        "Choose Excel file",

        type=[
            "xlsx",
            "xls",
        ],

        key="daily_upload",

    )


    if uploaded_file is not None:

        st.write(
            f"**Selected:** {uploaded_file.name}"
        )

        st.divider()


        if st.button(

            "🚀 Process & Update Dashboard",

            use_container_width=True,

            type="primary",

        ):

            try:

                # ==================================================
                # SAVE ORIGINAL UPLOAD
                # ==================================================

                upload_destination = (

                    UPLOAD_DIR
                    / uploaded_file.name

                )

                with open(
                    upload_destination,
                    "wb",
                ) as f:

                    f.write(
                        uploaded_file.getbuffer()
                    )

                st.success(
                    "Excel uploaded successfully."
                )


                # ==================================================
                # COPY TO PROGRESS DIRECTORY
                # ==================================================

                progress_destination = (

                    PROGRESS_DIR
                    / uploaded_file.name

                )

                shutil.copy2(

                    upload_destination,

                    progress_destination,

                )

                st.info(
                    "Daily report copied to the progress folder."
                )


                # ==================================================
                # RUN EXISTING PROGRESS SERVICE
                # ==================================================

                with st.spinner(
                    "Processing daily Excel..."
                ):

                    result = run_progress_service()


                # ==================================================
                # PROCESS RESULT
                # ==================================================

                if result.returncode != 0:

                    st.error(
                        "Progress processing failed."
                    )

                    if result.stdout:

                        st.text_area(

                            "Progress Service Output",

                            result.stdout,

                            height=250,

                        )

                    if result.stderr:

                        st.text_area(

                            "Progress Service Error",

                            result.stderr,

                            height=300,

                        )

                    st.stop()


                # ==================================================
                # SUCCESS
                # ==================================================

                st.success(
                    "Progress processed successfully."
                )

                if result.stdout:

                    with st.expander(
                        "View processing details"
                    ):

                        st.code(
                            result.stdout
                        )


                # ==================================================
                # CLEAR STREAMLIT CACHE
                # ==================================================

                st.cache_data.clear()

                st.cache_resource.clear()

                time.sleep(
                    1
                )


                # ==================================================
                # RELOAD PAGE
                # ==================================================

                st.rerun()


            except Exception as exc:

                st.error(
                    f"Upload failed: {exc}"
                )


    # ============================================================
    # SIDEBAR DASHBOARD INFORMATION
    # ============================================================

    st.divider()

    st.header(
        "📊 Dashboard"
    )

    st.write(
        "All dashboard calculations are generated "
        "directly by DashboardService."
    )

    st.divider()

    st.header(
        "Structure"
    )

    st.metric(
        "Sectors",
        EXPECTED_SECTORS,
    )

    st.metric(
        "Cells",
        EXPECTED_CELLS,
    )

    st.metric(
        "Villages",
        EXPECTED_VILLAGES,
    )


# ============================================================
# LOAD DASHBOARD SERVICE
# ============================================================

try:

    (
        service,
        district,
        sectors,
        cells,
        villages,
    ) = load_dashboard()


except Exception as exc:

    st.error(
        "Unable to load dashboard."
    )

    st.exception(
        exc
    )

    st.stop()


# ============================================================
# DATES
# ============================================================

previous_date = district[
    "previous_date"
]

current_date = district[
    "current_date"
]


previous_date_text = (

    pd.Timestamp(
        previous_date
    ).strftime(
        "%d/%m/%Y"
    )

)


current_date_text = (

    pd.Timestamp(
        current_date
    ).strftime(
        "%d/%m/%Y"
    )

)


# ============================================================
# DATE COMPARISON
# ============================================================

st.markdown(
    '<div class="section-title">📅 Comparison Period</div>',
    unsafe_allow_html=True,
)

date_col1, date_col2 = st.columns(
    2
)

with date_col1:

    st.info(
        f"Previous Report: {previous_date_text}"
    )


with date_col2:

    st.success(
        f"Current Report: {current_date_text}"
    )


# ============================================================
# EXTRACT DISTRICT VALUES
# ============================================================

baseline_money = district[
    "baseline_amount"
]

total_paid = district[
    "current_amount"
]

previous_paid = district[
    "previous_amount"
]

today_collected = district[
    "today_amount"
]

remaining_money = district[
    "remaining_amount"
]


baseline_members = district[
    "baseline_members"
]

total_members = district[
    "current_members"
]

previous_members = district[
    "previous_members"
]

today_enrolled = district[
    "today_members"
]

remaining_members = district[
    "remaining_members"
]


money_progress = district[
    "money_progress_percent"
]

member_progress = district[
    "member_progress_percent"
]


# ============================================================
# MONEY PERFORMANCE
# ============================================================

st.markdown(
    '<div class="section-title">💰 Money Performance</div>',
    unsafe_allow_html=True
)


c1, c2, c3, c4, c5 = st.columns(
    5
)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Baseline</div>
            <div class="kpi-value">
                {fmt_money(baseline_money)}
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
                Total Paid — {current_date_text}
            </div>
            <div class="kpi-value">
                {fmt_money(total_paid)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Previous Paid — {previous_date_text}
            </div>
            <div class="kpi-value">
                {fmt_money(previous_paid)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Today Collected
            </div>
            <div class="kpi-value">
                {fmt_money(today_collected)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c5:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Remaining
            </div>
            <div class="kpi-value">
                {fmt_money(remaining_money)}
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


c1, c2, c3, c4, c5 = st.columns(
    5
)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Baseline Members
            </div>
            <div class="kpi-value">
                {fmt_number(baseline_members)}
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
                Total — {current_date_text}
            </div>
            <div class="kpi-value">
                {fmt_number(total_members)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Previous — {previous_date_text}
            </div>
            <div class="kpi-value">
                {fmt_number(previous_members)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Today Enrolled
            </div>
            <div class="kpi-value">
                {fmt_number(today_enrolled)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c5:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                Remaining Members
            </div>
            <div class="kpi-value">
                {fmt_number(remaining_members)}
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


c1, c2 = st.columns(
    2
)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">
                💰 Money Progress
            </div>
            <div class="kpi-value">
                {fmt_percent(money_progress)}
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
            <div class="kpi-value">
                {fmt_percent(member_progress)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# DAILY SUMMARY
# ============================================================

st.markdown(
    '<div class="section-title">📊 Daily Summary</div>',
    unsafe_allow_html=True,
)


summary1, summary2, summary3, summary4 = st.columns(
    4
)


with summary1:

    st.metric(
        "Today Enrolled",
        fmt_number(
            today_enrolled
        ),
    )


with summary2:

    st.metric(
        "Today Collected",
        fmt_money(
            today_collected
        ),
    )


with summary3:

    st.metric(
        "Member Progress",
        fmt_percent(
            member_progress
        ),
    )


with summary4:

    st.metric(
        "Money Progress",
        fmt_percent(
            money_progress
        ),
    )


# ============================================================
# SECTOR PERFORMANCE
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">🥇 Sector Ranking</div>',
    unsafe_allow_html=True,
)


sector_table = prepare_table(
    sectors,
    "sector",
)


st.dataframe(

    sector_table,

    use_container_width=True,

    hide_index=True,

    height=550,

)


# ============================================================
# CELL PERFORMANCE
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">🥈 Cell Ranking</div>',
    unsafe_allow_html=True,
)


cell_table = prepare_table(
    cells,
    "cell",
)


st.dataframe(

    cell_table,

    use_container_width=True,

    hide_index=True,

    height=550,

)


# ============================================================
# VILLAGE PERFORMANCE
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">🥉 Village Ranking</div>',
    unsafe_allow_html=True,
)


village_table = prepare_table(
    villages,
    "village",
)


st.dataframe(

    village_table,

    use_container_width=True,

    hide_index=True,

    height=600,

)


# ============================================================
# DATA VALIDATION
# ============================================================

st.divider()

st.markdown(
    '<div class="section-title">✅ Data Validation</div>',
    unsafe_allow_html=True,
)


try:

    validation = (
        service.validate_structure()
    )

    v1, v2, v3 = st.columns(
        3
    )


    with v1:

        st.metric(

            "Sectors",

            validation[
                "progress"
            ][
                "sectors"
            ],

        )


    with v2:

        st.metric(

            "Cells",

            validation[
                "progress"
            ][
                "cells"
            ],

        )


    with v3:

        st.metric(

            "Villages",

            validation[
                "progress"
            ][
                "villages"
            ],

        )


except Exception as exc:

    st.warning(
        f"Validation unavailable: {exc}"
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Rubavu MUSA Monitoring Dashboard | "
    "Calculations are generated by DashboardService"
)

st.caption(
    f"Current report: {current_date_text} | "
    f"Previous report: {previous_date_text}"
)