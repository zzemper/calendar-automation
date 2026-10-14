# -*- coding: utf-8 -*-
"""
Created on Tue Oct 13 14:33:09 2026

@author: 3354745
"""

import io
import re
from datetime import date, datetime, time, timedelta

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE SETUP
# ============================================================

st.set_page_config(
    page_title="Calendar Automation",
    page_icon="📅",
    layout="wide",
)

st.title("📅 Calendar Automation")
st.write(
    "Select fiscal weeks and calendars, then click each date to cycle "
    "between Not selected, Open, and Closed."
)


# ============================================================
# CONSTANTS
# ============================================================

BASE_FISCAL_WEEK = 1
BASE_FISCAL_WEEK_START = date(2026, 2, 1)  # Sunday
MAX_FISCAL_WEEK = 52

AVAILABLE_CALENDARS = [
    "SHIP_DF_860_TG",
    "New calendar",
]

MINUTE_OPTIONS = [
    "First minute of the day",
    "Last minute of the day",
    "2nd to last minute of the day",
]

DAY_STATUS_SEQUENCE = ["Not selected", "Open", "Closed"]

STATUS_STYLE = {
    "Not selected": {
        "icon": "⚪",
        "label": "",
    },
    "Open": {
        "icon": "🟢",
        "label": "OPEN",
    },
    "Closed": {
        "icon": "🔴",
        "label": "CLOSED",
    },
}

WEEKDAY_NAMES = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
]

OUTPUT_COLUMNS = [
    "*Cal",
    "*Eff",
    "AllocWgt",
    "AltCal",
    "Avail",
    "Descr",
    "Opt",
    "PerWgt",
    "Repeat",
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def safe_key(value):
    """Convert text into a safe Streamlit widget/session key."""
    return re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_")


def fiscal_week_start(week_number):
    """Return the Sunday starting the requested 2026 fiscal week."""
    return BASE_FISCAL_WEEK_START + timedelta(
        weeks=week_number - BASE_FISCAL_WEEK
    )


def dates_for_fiscal_week(week_number):
    """Return the seven dates in a fiscal week, Sunday through Saturday."""
    start_date = fiscal_week_start(week_number)
    return [start_date + timedelta(days=offset) for offset in range(7)]


def date_status_key(calendar_name, calendar_date):
    """Create a unique state key for a calendar and date."""
    return (
        f"day_status__{safe_key(calendar_name)}__"
        f"{calendar_date.isoformat()}"
    )


def get_day_status(calendar_name, calendar_date):
    """Return a date's current state, defaulting to gray/not selected."""
    key = date_status_key(calendar_name, calendar_date)
    return st.session_state.get(key, "Not selected")


def cycle_day_status(calendar_name, calendar_date):
    """Cycle gray -> green/open -> red/closed -> gray."""
    key = date_status_key(calendar_name, calendar_date)
    current_status = st.session_state.get(key, "Not selected")
    current_index = DAY_STATUS_SEQUENCE.index(current_status)
    next_index = (current_index + 1) % len(DAY_STATUS_SEQUENCE)
    st.session_state[key] = DAY_STATUS_SEQUENCE[next_index]


def clear_calendar_dates(calendar_name, fiscal_weeks):
    """Reset all visible dates for one calendar to gray/not selected."""
    for week_number in fiscal_weeks:
        for calendar_date in dates_for_fiscal_week(week_number):
            key = date_status_key(calendar_name, calendar_date)
            st.session_state[key] = "Not selected"


def mark_all_calendar_dates(calendar_name, fiscal_weeks, status):
    """Mark all visible dates for one calendar Open or Closed."""
    for week_number in fiscal_weeks:
        for calendar_date in dates_for_fiscal_week(week_number):
            key = date_status_key(calendar_name, calendar_date)
            st.session_state[key] = status


def build_effective_datetimes(calendar_date, minute_option):
    """Return the two *Eff values required for one selected date."""
    if minute_option == "First minute of the day":
        first_record = datetime.combine(calendar_date, time(0, 0))
        second_record = datetime.combine(calendar_date, time(0, 1))

    elif minute_option == "Last minute of the day":
        first_record = datetime.combine(calendar_date, time(23, 59))
        second_record = datetime.combine(
            calendar_date + timedelta(days=1),
            time(0, 0),
        )

    else:  # 2nd to last minute of the day
        first_record = datetime.combine(calendar_date, time(23, 58))
        second_record = datetime.combine(calendar_date, time(23, 59))

    return first_record, second_record


def build_output_dataframe(
    selected_calendars,
    fiscal_weeks,
    minute_selection_by_calendar,
    description,
    add_repeat,
):
    """
    Build the final import-format dataframe.

    Export rules:
    - Gray/Not selected dates are excluded.
    - Open dates create two rows with Opt = 2.
    - Closed dates create two rows with Opt = 1.
    - If add_repeat is enabled, one additional row is created
      for the day immediately after the latest selected date.
    - The additional repeat row receives Opt = 5 and Repeat = 7D.
    """

    rows = []

    for calendar_name in selected_calendars:
        minute_option = minute_selection_by_calendar[
            calendar_name
        ]

        calendar_rows = []
        selected_calendar_dates = []

        for week_number in fiscal_weeks:
            for calendar_date in dates_for_fiscal_week(
                week_number
            ):
                status = get_day_status(
                    calendar_name,
                    calendar_date
                )

                # Gray dates are intentionally excluded.
                if status == "Not selected":
                    continue

                # Track the actual calendar date selected by the user.
                selected_calendar_dates.append(
                    calendar_date
                )

                if status == "Open":
                    opt_value = 2
                else:
                    opt_value = 1

                first_eff, second_eff = (
                    build_effective_datetimes(
                        calendar_date,
                        minute_option
                    )
                )

                # First record for the selected date.
                calendar_rows.append(
                    {
                        "*Cal": calendar_name,
                        "*Eff": first_eff,
                        "AllocWgt": 0,
                        "AltCal": "",
                        "Avail": 0,
                        "Descr": description.strip(),
                        "Opt": opt_value,
                        "PerWgt": 0,
                        "Repeat": "",
                    }
                )

                # Second record for the selected date.
                calendar_rows.append(
                    {
                        "*Cal": calendar_name,
                        "*Eff": second_eff,
                        "AllocWgt": 0,
                        "AltCal": "",
                        "Avail": 0,
                        "Descr": description.strip(),
                        "Opt": opt_value,
                        "PerWgt": 0,
                        "Repeat": "",
                    }
                )

        # ----------------------------------------------------
        # ADD THE NEW REPEAT ROW
        # ----------------------------------------------------

        if (
            add_repeat
            and selected_calendar_dates
        ):
            latest_selected_date = max(
                selected_calendar_dates
            )

            repeat_date = (
                latest_selected_date
                + timedelta(days=1)
            )

            if minute_option == "First minute of the day":
                repeat_effective_datetime = (
                    datetime.combine(
                        repeat_date,
                        time(0, 0)
                    )
                )

            elif minute_option == "Last minute of the day":
                repeat_effective_datetime = (
                    datetime.combine(
                        repeat_date,
                        time(23, 59)
                    )
                )

            else:
                # 2nd to last minute of the day
                repeat_effective_datetime = (
                    datetime.combine(
                        repeat_date,
                        time(23, 58)
                    )
                )

            calendar_rows.append(
                {
                    "*Cal": calendar_name,
                    "*Eff": repeat_effective_datetime,
                    "AllocWgt": 0,
                    "AltCal": "",
                    "Avail": 0,
                    "Descr": description.strip(),
                    "Opt": 5,
                    "PerWgt": 0,
                    "Repeat": "7D",
                }
            )

        calendar_rows.sort(
            key=lambda row: row["*Eff"]
        )

        rows.extend(calendar_rows)

    output_df = pd.DataFrame(
        rows,
        columns=OUTPUT_COLUMNS
    )

    if not output_df.empty:
        output_df["*Eff"] = pd.to_datetime(
            output_df["*Eff"]
        )

        output_df = output_df.sort_values(
            by=["*Cal", "*Eff"],
            kind="stable"
        ).reset_index(drop=True)

    return output_df


def create_excel_file(output_df):
    """Create the formatted Excel import file in memory."""
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        datetime_format="m/d/yyyy h:mm",
    ) as writer:
        output_df.to_excel(
            writer,
            index=False,
            sheet_name="Calendar Import",
        )

        worksheet = writer.book["Calendar Import"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )
        header_font = Font(
            color="FFFFFF",
            bold=True,
        )

        for header_cell in worksheet[1]:
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for date_cell in worksheet["B"][1:]:
            date_cell.number_format = "m/d/yyyy h:mm"

        widths = {
            "A": 26,
            "B": 20,
            "C": 12,
            "D": 12,
            "E": 10,
            "F": 30,
            "G": 8,
            "H": 12,
            "I": 10,
        }

        for column_letter, width in widths.items():
            worksheet.column_dimensions[column_letter].width = width

        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="center")

    output.seek(0)
    return output


def selected_date_summary(calendar_name, fiscal_weeks):
    """Count the selected open, closed, and gray dates."""
    counts = {
        "Open": 0,
        "Closed": 0,
        "Not selected": 0,
    }

    for week_number in fiscal_weeks:
        for calendar_date in dates_for_fiscal_week(week_number):
            status = get_day_status(calendar_name, calendar_date)
            counts[status] += 1

    return counts


# ============================================================
# STEP 1: FISCAL WEEKS
# ============================================================

st.header("Step 1: Select fiscal weeks")

st.info(
    "Fiscal week 40 begins Sunday, November 1, 2026. "
    "Each following fiscal week begins seven days later."
)

week_col1, week_col2 = st.columns(2)

with week_col1:
    starting_week = st.number_input(
        "Starting fiscal week",
        min_value=BASE_FISCAL_WEEK,
        max_value=MAX_FISCAL_WEEK,
        value=BASE_FISCAL_WEEK,
        step=1,
    )

with week_col2:
    ending_week = st.number_input(
        "Ending fiscal week",
        min_value=BASE_FISCAL_WEEK,
        max_value=MAX_FISCAL_WEEK,
        value=47,
        step=1,
    )

if ending_week < starting_week:
    st.error("Ending fiscal week must be the same as or later than the starting week.")
    st.stop()

fiscal_weeks = list(range(int(starting_week), int(ending_week) + 1))

week_reference = pd.DataFrame(
    [
        {
            "Fiscal Week": week_number,
            "Start Date": fiscal_week_start(week_number),
            "End Date": fiscal_week_start(week_number) + timedelta(days=6),
        }
        for week_number in fiscal_weeks
    ]
)

st.dataframe(
    week_reference,
    hide_index=True,
    use_container_width=True,
)


# ============================================================
# STEP 2: CALENDAR SELECTION
# ============================================================

st.header("Step 2: Select calendars to update")

existing_calendars = st.multiselect(
    "Existing calendars",
    options=[
        "SHIP_DF_860_TG",
        "ARRIVE_CENTFILL_RX",
        "ARRIVE_CF_RX464",
        "ARRIVE_DF805_881_HOLIDAY",
        "ARRIVE_DF805_FOODS",
        "ARRIVE_DF807_FOODS",
        "ARRIVE_DF84_FOODS",
        "ARRIVE_DF84_WED",
        "ARRIVE_DF86_ART",
        "ARRIVE_DF86_TUWTH",
        "ARRIVE_DF871_FOODS",
        "ARRIVE_DF871_SU-F",
        "ARRIVE_DF871_THU",
        "ARRIVE_DF_807_MON_CLOSED",
        "ARRIVE_DF_86_GAS",
        "ARRIVE_DF_86_THUR_LASTMIN ",
        "ARRIVE_DF_888",
        "ARRIVE_DF_ART_THUR_LASTMIN",
        "ARRIVE_DF_CIGS",
        "ARRIVE_DF_FOODS",
        "ARRIVE_DF_FOODS_86",
        "ARRIVE_DF_FRI_FIRSTMIN",
        "ARRIVE_DF_FRI_LASTMIN",
        "ARRIVE_DF_FZICE",
        "ARRIVE_DF_GM",
        "ARRIVE_DF_GM_800",
        "ARRIVE_DF_GM_802",
        "ARRIVE_DF_GM_86_LASTMIN",
        "ARRIVE_DF_GM_90",
        "ARRIVE_DF_GM_93",
        "ARRIVE_DF_GM_NOMON",
        "ARRIVE_DF_GM_TUES",
        "ARRIVE_DF_GM_TUES_THUR",
        "ARRIVE_DF_GM_WED_FIRSTMIN",
        "ARRIVE_DF_GM_WED_LASTMIN",
        "ARRIVE_DF_GM_WR",
        "ARRIVE_DF_M-F_LASTMIN",
        "ARRIVE_DF_MTWRF_FIRSTMIN",
        "ARRIVE_DF_MTWR_LASTMIN",
        "ARRIVE_DF_MWRFS_LASTMIN",
        "ARRIVE_DF_MWRF_LASTMIN",
        "ARRIVE_DF_MWTHF_805",
        "ARRIVE_DF_RX_90",
        "ARRIVE_DF_SUN_THURS_LASTMIN",
        "ARRIVE_DF_THUR_LASTMIN",
        "ARRIVE_DF_THUR_LASTMIN ",
        "ARRIVE_DF_TUE_LASTMIN",
        "ARRIVE_DF_TWR_LASTMIN",
        "ARRIVE_DF_WED_LASTMIN",
        "ARRIVE_STORE_860_FRZHOLIDAY",
        "ARRIVE_STORE_ART871",
        "ARRIVE_STORE_CIGS",
        "ARRIVE_STORE_FOODS",
        "ARRIVE_STORE_FOODS_12_26",
        "ARRIVE_STORE_FOODS_2",
        "ARRIVE_STORE_FOODS_FRESH_12_26",
        "ARRIVE_STORE_FRESH_SEAFOOD",
        "ARRIVE_STORE_GARDEN",
        "ARRIVE_STORE_GM",
        "ARRIVE_STORE_GM_NOSUN2",
        "ARRIVE_STORE_GM_NOSUNDAY",
        "ARRIVE_STORE_GM_TEST",
        "ARRIVE_STORE_NEXTDAYDEMAND",
        "ARRIVE_STORE_RX",
        "ARRIVE_STORE_RX_465",
        "ARRIVE_STORE_RX_466",
        "ARRIVE_STORE_RX_TWHF",
        "ARRIVE_STORE_SAMEDAYDEMAND",
        "ARRIVE_STORE_SAMEDAY_NO_WED",
        "ARRIVE_STORE_SUPPLIES",
        "ARRIVE_STR_UNFI",
        "ARRIVE_STR_UNFI_B",
        "ARRIV_DF_FRI_LAST_MILK",
        "ARR_7D_FRESH",
        "ARR_7D_HIGHLINE",
        "ARR_860_FRSH_MWF",
        "ARR_860_FRSH_WED_ALLDC",
        "ARR_860_FRSH_WED_NEXTDAY",
        "ARR_860_THU_FRSH_ALLDC_NEXTDAY",
        "ARR_860_THU_FRZN_ALLDC_NEXTDAY",
        "ARR_ALLDAY_FRESH",
        "ARR_DF882_FRZN_TU-SA",
        "ARR_DF_7D_FIRSTMILK",
        "ARR_DF_7D_LASTMILK",
        "ARR_DF_7D_NEXTDAY",
        "ARR_DF_801_TCD",
        "ARR_DF_801_TCD_7DAY",
        "ARR_DF_801_TCD_F",
        "ARR_DF_801_TCD_M",
        "ARR_DF_801_TCD_MR",
        "ARR_DF_801_TCD_MRF",
        "ARR_DF_801_TCD_MTWR",
        "ARR_DF_801_TCD_MWRF",
        "ARR_DF_801_TCD_TU",
        "ARR_DF_860_FRESH_M-FR",
        "ARR_DF_860_FRZN_M-FR",
        "ARR_DF_883_TCD_DF-DF",
        "ARR_DF_883_TCD_DF-DF_TF",
        "ARR_DF_883_TCD_DF-DF_THU",
        "ARR_DF_CLOSURES1",
        "ARR_DF_CLOSURES2",
        "ARR_DF_DELI_TUTHFRI_47",
        "ARR_DF_DRYCHS_M-SA",
        "ARR_DF_FIRSTMILK_SuWF",
        "ARR_DF_FIRSTMIN_FRI",
        "ARR_DF_FIRSTMIN_MONMILK",
        "ARR_DF_FIRSTMIN_MWF",
        "ARR_DF_FIRSTMIN_SAT",
        "ARR_DF_FIRSTMIN_SUN",
        "ARR_DF_FIRSTMIN_SUWF",
        "ARR_DF_FIRSTMIN_THUR",
        "ARR_DF_FIRSTMIN_THURMILK",
        "ARR_DF_FIRSTMIN_TRSA ",
        "ARR_DF_FIRSTMIN_TTHSA",
        "ARR_DF_FIRSTMIN_WSA",
        "ARR_DF_FRESH_MWRFSA",
        "ARR_DF_FRESH_MWRFSA_NEXTDAY",
        "ARR_DF_FRESH_TU-SA",
        "ARR_DF_FRI_FIRSTMILK",
        "ARR_DF_FRI_FRZN",
        "ARR_DF_FRSH_ALLDC_M-Sa",
        "ARR_DF_FRSH_ALLDC_MSa",
        "ARR_DF_FRSH_ALLDC_SuF",
        "ARR_DF_FRSH_ALLDC_SuTWRFSa",
        "ARR_DF_FRSH_ALLDC_TF",
        "ARR_DF_FRSH_ALLDC_W-Sa",
        "ARR_DF_FRSH_MILK_1STMIN",
        "ARR_DF_FRZN_89_7DAY",
        "ARR_DF_FRZN_89_M-SA",
        "ARR_DF_FRZN_89_MWRSA",
        "ARR_DF_FRZN_89_SU-F",
        "ARR_DF_FRZN_89_SuTR",
        "ARR_DF_FRZN_89_WFS",
        "ARR_DF_FRZN_89_WSA",
        "ARR_DF_FRZN_89_WSU",
        "ARR_DF_FRZN_M-FR",
        "ARR_DF_FRZN_M-SA",
        "ARR_DF_FRZN_M-T-R",
        "ARR_DF_FRZN_MON",
        "ARR_DF_FRZN_MON-SAT",
        "ARR_DF_FRZN_MRSASU",
        "ARR_DF_FRZN_MT",
        "ARR_DF_FRZN_MTRF",
        "ARR_DF_FRZN_MWR",
        "ARR_DF_FRZN_RF",
        "ARR_DF_FRZN_SAT",
        "ARR_DF_FRZN_SFD_7D",
        "ARR_DF_FRZN_SMTRF",
        "ARR_DF_FRZN_SU-TU",
        "ARR_DF_FRZN_SUN",
        "ARR_DF_FRZN_SUN_TUE",
        "ARR_DF_FRZN_SUN_WED",
        "ARR_DF_FRZN_SUN_WED_FRI",
        "ARR_DF_FRZN_SuTh",
        "ARR_DF_FRZN_SUTRFSA ",
        "ARR_DF_FRZN_SUTRSA",
        "ARR_DF_FRZN_SU_TU_SA",
        "ARR_DF_FRZN_TRSASU",
        "ARR_DF_FRZN_TU-SA",
        "ARR_DF_FRZN_TUE",
        "ARR_DF_FRZN_TUE-SAT",
        "ARR_DF_FRZN_TUE_FRI",
        "ARR_DF_FRZN_TUE_THU",
        "ARR_DF_FRZN_TU_TH",
        "ARR_DF_FRZN_UMWRFS",
        "ARR_DF_FRZN_UMWRS",
        "ARR_DF_FRZN_WED-SUN",
        "ARR_DF_GAS_FRESH",
        "ARR_DF_GM_86_MWF",
        "ARR_DF_GM_86_SUTUTH",
        "ARR_DF_GROC_MON_LASTMIN ",
        "ARR_DF_GROC_SUN_LASTMIN",
        "ARR_DF_JUICE_7D",
        "ARR_DF_LASTMILK_SAT",
        "ARR_DF_LASTMILK_SATWED",
        "ARR_DF_LASTMILK_SUN",
        "ARR_DF_LAST_MTHSA_MILK",
        "ARR_DF_LAST_MTWR_MILK ",
        "ARR_DF_LAST_MWFSA",
        "ARR_DF_LAST_MWRF_MILK",
        "ARR_DF_LAST_SUMWRF",
        "ARR_DF_LAST_SUN_MILK ",
        "ARR_DF_M-FR_860_FRSH_NEXTDAY",
        "ARR_DF_M-FR_860_FRZN_NEXTDAY",
        "ARR_DF_M-FR_FRZN_NEXTDAY",
        "ARR_DF_M-FR_NEXTDAY",
        "ARR_DF_M-SA_NEXTDAY",
        "ARR_DF_MEAT_7D",
        "ARR_DF_MEAT_STRF",
        "ARR_DF_MID_M-FR",
        "ARR_DF_MID_T-SU",
        "ARR_DF_MTWRSa_NEXT",
        "ARR_DF_MTWRS_FIRSTMIN",
        "ARR_DF_MTWR_MILKLAST",
        "ARR_DF_MWHFS_NEXTDAY",
        "ARR_DF_PMT_SFD_M-SA",
        "ARR_DF_PRODUCE_7D",
        "ARR_DF_RF_FRZN",
        "ARR_DF_SFD_7D",
        "ARR_DF_SuMTWRSa_NEXT",
        "ARR_DF_SuMTWR_LAST",
        "ARR_DF_SUTWRF_ALLDC",
        "ARR_DF_SWRF_FRSH",
        "ARR_DF_S_T_F_ALLDC",
        "ARR_DF_T-SU_NEXTDAY",
        "ARR_DF_THUR_LASTMILK",
        "ARR_DF_THU_FIRSTMIN",
        "ARR_DF_TR_FRSH_NEXTDAY",
        "ARR_DF_TR_NEXTDAY",
        "ARR_DF_TWRF_FIRSTMIN",
        "ARR_DF_WED_FIRSTMIN",
        "ARR_DRY_WED_NEXTDAY",
        "ARR_DRY_WF",
        "ARR_FRESH_MRS",
        "ARR_FRESH_MTRS",
        "ARR_FRESH_STWRFS",
        "ARR_FRESH_T_R_F",
        "ARR_FRI_ALLDC",
        "ARR_FRI_ALLDC_NEXTDAY",
        "ARR_FRI_FRSH_ALLDC",
        "ARR_FRI_PLEX_DC",
        "ARR_FRSH_MF_ALLDC",
        "ARR_FRSH_MON_ALLDC",
        "ARR_FRSH_MTWRF_ALLDC",
        "ARR_FRSH_MTWR_ALLDC",
        "ARR_FRSH_MWF",
        "ARR_FRSH_MW_ALLDC",
        "ARR_FRSH_RFS",
        "ARR_FRSH_TF",
        "ARR_FRSH_TRFS ",
        "ARR_FRSH_TWF",
        "ARR_FRSH_TWFS",
        "ARR_FRSH_TWRFSa",
        "ARR_FRSH_WED_ALLDC",
        "ARR_FRSH_WED_ALLDC_SAMEDAY",
        "ARR_FRSH_WF",
        "ARR_FRSH_WFS",
        "ARR_FRSH_WRF",
        "ARR_FRSH_WRSa",
        "ARR_FRZN_MIN_WED_ALLDC",
        "ARR_FRZN_MON_WED_ALLDC",
        "ARR_FRZN_MON_WED_ALLDC ",
        "ARR_FRZN_MR_NEXTDAY",
        "ARR_FRZN_MW_ALLDC",
        "ARR_FRZN_WED_ALLDC",
        "ARR_FRZN_WED_NEXTDAY",
        "ARR_FRZN_WR_ALLDC",
        "ARR_FRZ_WFS",
        "ARR_FSU_FRSH_ALLDC",
        "ARR_LAST_TH_FR_SA",
        "ARR_M-F_ALLDC",
        "ARR_M-F_TCD_801",
        "ARR_M-S_ALLDC",
        "ARR_MF_ALLDC",
        "ARR_MON_881",
        "ARR_MON_ALLDC",
        "ARR_MON_ALLDC_NEXTDAY",
        "ARR_MON_FRI_FRSH",
        "ARR_MON_FRZN_DF",
        "ARR_MO_SA_PEARSON",
        "ARR_MRF_SANDWICH",
        "ARR_MRS_ALLDC",
        "ARR_MR_ALLDC",
        "ARR_MR_ALLDC_NEXTDAY",
        "ARR_MR_FRSH_ALL",
        "ARR_MR_FRSH_ALLDC",
        "ARR_MR_FRSH_ALLDC_NEXTDAY",
        "ARR_MR_FRSH_TCD",
        "ARR_MSA_FRSH_ALLDC",
        "ARR_MSA_FRSH_ALLDC_NEXTDAY",
        "ARR_MTHS_ALLDC",
        "ARR_MTRFS_ALLDC",
        "ARR_MTRS_ALLDC",
        "ARR_MTR_FRSH_ALLDC",
        "ARR_MTR_FRSH_ALLDC_NEXTDAY",
        "ARR_MTWF_ALLDC_NEXTDAY",
        "ARR_MT_ALLDC_NEXTDAY",
        "ARR_MT_FRSH_ALLDC",
        "ARR_MWFS_ALLDC",
        "ARR_MWF_ALLDC",
        "ARR_MWF_FRSH_ALL",
        "ARR_MWRFS_ALLDC",
        "ARR_MWRF_ALLDC",
        "ARR_MWRSASU_ALLDC",
        "ARR_MWR_ALLDC",
        "ARR_MWR_ALLDC_NEXTDAY",
        "ARR_MWR_FRZN_ALLDC_NEXTDAY",
        "ARR_MWS_ALLDC",
        "ARR_MW_ALLDC",
        "ARR_RF_ALLDC",
        "ARR_RF_ALLDC_NEXTDAY",
        "ARR_RSA_FRSH_ALLDC",
        "ARR_S-F_ALLDC",
        "ARR_SAT_ALLDC",
        "ARR_SAT_ALLDC_NEXTDAY",
        "ARR_SAT_FRESH",
        "ARR_SAT_FRSH_ALLDC",
        "ARR_SMTHS_ALLDC",
        "ARR_SMTWRSa_FRSH",
        "ARR_SMWRF_ALLDC",
        "ARR_SR_ALLDC",
        "ARR_STR271_PCC_DFD",
        "ARR_STR_ALLDC",
        "ARR_SUN_ALLDC",
        "ARR_SUTWRF_HIGHLINE",
        "ARR_SU_MO_FR_ALLDC_NEXTDAY",
        "ARR_SWF_ALLDC",
        "ARR_SW_ALLDC",
        "ARR_T-M_ALLDC",
        "ARR_TF_ALLDC",
        "ARR_TF_ALLDC_NEXTDAY",
        "ARR_THURS_FRESH",
        "ARR_THU_860_FRZN_ALLDC",
        "ARR_THU_ALLDC",
        "ARR_THU_ALLDC_NEXTDAY",
        "ARR_THU_DF_86_GAS",
        "ARR_THU_DRY_ALLDC",
        "ARR_THU_FRSH_ALLDC",
        "ARR_THU_FRSH_ALLDC_HOL_FRI",
        "ARR_THU_FRZN_ALLDC",
        "ARR_THU_FRZN_ALLDC_NEXTDAY",
        "ARR_THU_TCD_801",
        "ARR_TH_FRZN_ALLDC_NEXTDAY",
        "ARR_TRF_SANDWICH",
        "ARR_TRS_ALLDC",
        "ARR_TR_ALLDC",
        "ARR_TU-FR_860_FRZN_ALLDC_NEXTDAY",
        "ARR_TU-FR_ALLDC",
        "ARR_TU-FR_ALLDC_NEXTDAY",
        "ARR_TU-FR_FRZN_ALLDC_NEXTDAY",
        "ARR_TUE_ALLDC",
        "ARR_TUE_ALLDC_NEXTDAY",
        "ARR_TUE_THU_FRSH_ALLDC",
        "ARR_TUWED_FRSH",
        "ARR_TU_86_805",
        "ARR_TU_FRSH_ALLDC",
        "ARR_TU_SAT_FRSH",
        "ARR_TU_THU_FRSH",
        "ARR_TU_TH_TCD_801",
        "ARR_TWF_ALLDC_NEXTDAY",
        "ARR_TWRFSS_PEARSON",
        "ARR_TWR_ALLDC",
        "ARR_T_R_SA_PEARSON",
        "ARR_W-FR_FRZN_ALLDC_NEXTDAY",
        "ARR_W-FR_FRZN_ALLDC_NEXTDAY ",
        "ARR_WED_ALLDC",
        "ARR_WED_ALLDC_NEXTDAY",
        "ARR_WED_DF_86_GAS",
        "ARR_WED_FRI_FRSH",
        "ARR_WED_SAT_FRSH",
        "ARR_WF_ALLDC",
        "ARR_WF_TCD_801",
        "ARR_WR_ALLDC",
        "ARR_WR_ALLDC_NEXTDAY",
        "ARR_WR_FRSH_ALLDC  ",
        "ARR_WS_ALLDC",
        "SHIP_860_TRANSITION 1",
        "SHIP_860_TRS_TRANS",
        "SHIP_871_BEFORE",
        "SHIP_CENTA_2ND",
        "SHIP_CENTB_2ND",
        "SHIP_CENTC_2ND",
        "SHIP_DF-DF_860_804",
        "SHIP_DF-DF_860_872",
        "SHIP_DF-DF_860_882",
        "SHIP_DF-DF_860_89",
        "SHIP_DF89_4X_2ND",
        "SHIP_DF89_5X_2ND",
        "SHIP_DF_801_7DAY",
        "SHIP_DF_801_NO_TU_PROD",
        "SHIP_DF_801_NO_TU_SAMEDAYDMND",
        "SHIP_DF_801_TCD",
        "SHIP_DF_801_TCD_DF-DF_MWF",
        "SHIP_DF_801_TCD_DF-DF_TU_TH_SA",
        "SHIP_DF_801_TCD_DF-DF_WSU",
        "SHIP_DF_801_TCD_M",
        "SHIP_DF_801_TCD_MTW",
        "SHIP_DF_801_TCD_MWRF",
        "SHIP_DF_801_TCD_RSA_872_LT2D",
        "SHIP_DF_801_TCD_TH_SA",
        "SHIP_DF_804_GRP_B",
        "SHIP_DF_804_GRP_B  _STOREORD",
        "SHIP_DF_804_GRP_B_SEAFOOD",
        "SHIP_DF_804_NO_TU_SAMEDAYDMND",
        "SHIP_DF_807",
        "SHIP_DF_807_MARKET",
        "SHIP_DF_807_MTFS_CNTRL_SRV",
        "SHIP_DF_807_MTRFS",
        "SHIP_DF_807_MTRF_CNTRL_SRV",
        "SHIP_DF_807_MTWFS",
        "SHIP_DF_807_MWFS",
        "SHIP_DF_807_MWF_CNTRL_SRV",
        "SHIP_DF_807_TEMP_HOLIDAY25_6XWK48",
        "SHIP_DF_807_TRFS_CNTRL_SRV",
        "SHIP_DF_807_TWFS_CNTRL_SRV",
        "SHIP_DF_84",
        "SHIP_DF_84 to 871 Stores",
        "SHIP_DF_84_A",
        "SHIP_DF_84_B",
        "SHIP_DF_84_MARKET",
        "SHIP_DF_84_MFS_CNTRL_SRV",
        "SHIP_DF_84_MFS_LVITEMS",
        "SHIP_DF_84_MTFS_CNTRL_SRV",
        "SHIP_DF_84_MTFS_LVITEMS",
        "SHIP_DF_84_MTRFS",
        "SHIP_DF_84_MTRF_CNTRL_SRV",
        "SHIP_DF_84_MTWFS",
        "SHIP_DF_84_MTWRF_CNTRL_SRV",
        "SHIP_DF_84_MWF",
        "SHIP_DF_84_MWFS",
        "SHIP_DF_84_MWFS_4XPILOT",
        "SHIP_DF_84_MWFS_CNTRL_SRV",
        "SHIP_DF_84_MWF_CNTRL_SRV",
        "SHIP_DF_84_MWRF_CNTRL_SRV",
        "SHIP_DF_84_MWTF",
        "SHIP_DF_84_MWTHFS",
        "SHIP_DF_84_RS465",
        "SHIP_DF_84_RXSTORES",
        "SHIP_DF_84_SHIFT",
        "SHIP_DF_84_SHIFT1",
        "SHIP_DF_84_SHIFT2 ",
        "SHIP_DF_84_STRF",
        "SHIP_DF_84_TIPPA",
        "SHIP_DF_84_TIPPB",
        "SHIP_DF_84_TIPP_CNTRL_SRVA",
        "SHIP_DF_84_TIPP_CNTRL_SRVB",
        "SHIP_DF_84_TR",
        "SHIP_DF_84_TRFS_CNTRL_SRV",
        "SHIP_DF_84_TRS_CNTRL_SRV",
        "SHIP_DF_84_TWFS_CNTRL_SRV",
        "SHIP_DF_85_6DAY PCC",
        "SHIP_DF_85_7DAY",
        "SHIP_DF_85_7D_PCC",
        "SHIP_DF_85_DIPS",
        "SHIP_DF_85_FRESHROOM",
        "SHIP_DF_85_JUICE",
        "SHIP_DF_85_MR",
        "SHIP_DF_85_NO_TU_PCC",
        "SHIP_DF_85_NO_TU_PROD",
        "SHIP_DF_85_NO_TU_SAMEDAYDMND",
        "SHIP_DF_85_SHAREABLES",
        "SHIP_DF_85_SU_TU_TH",
        "SHIP_DF_860_A",
        "SHIP_DF_860_B",
        "SHIP_DF_860_CK",
        "SHIP_DF_860_SWF",
        "SHIP_DF_860_TG",
        "SHIP_DF_860_Thanksgiving2024",
        "SHIP_DF_860_TRANSITION 2",
        "SHIP_DF_860_TRSA",
        "SHIP_DF_860_TURKEYS",
        "SHIP_DF_86TO84_MWF",
        "SHIP_DF_86_CENTSERV",
        "SHIP_DF_86_CIGS_872",
        "SHIP_DF_86_GAS",
        "SHIP_DF_86_GAS_3_DAY",
        "SHIP_DF_86_GAS_TEMP",
        "SHIP_DF_86_HD-MARKET_FMT",
        "SHIP_DF_86_MARKET_FMT",
        "SHIP_DF_86_SUPP",
        "SHIP_DF_871",
        "SHIP_DF_871_4DAY",
        "SHIP_DF_871_5DAY",
        "SHIP_DF_871_5DAY - TUES",
        "SHIP_DF_871_5DAY - WED",
        "SHIP_DF_871_5DAY-THU",
        "SHIP_DF_871_5DAY-TUE",
        "SHIP_DF_871_5D_THU_B",
        "SHIP_DF_871_5D_WED_B",
        "SHIP_DF_871_A",
        "SHIP_DF_871_B",
        "SHIP_DF_871_MFS_LVITEMS",
        "SHIP_DF_871_MTFS_LVITEMS",
        "SHIP_DF_871_TIPP_A",
        "SHIP_DF_871_TIPP_B",
        "SHIP_DF_872_7DAY",
        "SHIP_DF_872_MILK_PCC",
        "SHIP_DF_872_NO_TU_PROD",
        "SHIP_DF_872_NO_TU_SAMEDAYDMND",
        "SHIP_DF_872_STORD_FRSHSEAFD",
        "SHIP_DF_882_DIPS",
        "SHIP_DF_882_FRSHSEAFD",
        "SHIP_DF_882_FRSHSEAFD_RGLS",
        "SHIP_DF_882_JUICE",
        "SHIP_DF_882_NO_TU_ALLSTR",
        "SHIP_DF_882_NO_TU_SAMEDAYDMND",
        "SHIP_DF_882_SEAFOOD_4xStores",
        "SHIP_DF_883_7DAY_PROD",
        "SHIP_DF_883_NON_PRODUCE_MARKET",
        "SHIP_DF_883_NO_TU_PROD",
        "SHIP_DF_883_NO_TU_SAMEDAYDMND",
        "SHIP_DF_888",
        "SHIP_DF_888_SMTThSa",
        "SHIP_DF_89MANU_CENT_A",
        "SHIP_DF_89MANU_CENT_B",
        "SHIP_DF_89MANU_CENT_C",
        "SHIP_DF_89MANU_CENT_D",
        "SHIP_DF_89_4xA-Centralized",
        "SHIP_DF_89_4xA_Manu",
        "SHIP_DF_89_4xA_SO",
        "SHIP_DF_89_4xB_Manu",
        "SHIP_DF_89_4xB_SEA",
        "SHIP_DF_89_4xB_SO",
        "SHIP_DF_89_5xA_Manu",
        "SHIP_DF_89_5xA_SEA",
        "SHIP_DF_89_5xA_SO",
        "SHIP_DF_89_5xB_Manu",
        "SHIP_DF_89_5xB_SEA",
        "SHIP_DF_89_5xB_SO",
        "SHIP_DF_89_6xA_Manu",
        "SHIP_DF_89_6xA_SEA",
        "SHIP_DF_89_6xA_SO",
        "SHIP_DF_89_6xB_Manu",
        "SHIP_DF_89_6xB_SEA",
        "SHIP_DF_89_CENT_5xA_Manu",
        "SHIP_DF_89_Manu_Centralized",
        "SHIP_DF_89_TG_4x",
        "SHIP_DF_89_TG_5x",
        "SHIP_DF_89_TG_5xB",
        "SHIP_DF_90_SUPP",
        "SHIP_DF_93_MARKET_FMT",
        "SHIP_DF_93_SUPP",
        "SHIP_DF_ART_HD-MTUTH",
        "SHIP_DF_ART_HD-MTUWTHF",
        "SHIP_DF_ART_HD-MWF",
        "SHIP_DF_ART_HD-SUTUTH",
        "SHIP_DF_ART_HD-SUTUW",
        "SHIP_DF_ART_HD-SUW",
        "SHIP_DF_ART_HD-TUTH",
        "SHIP_DF_ART_HOLCAL_1",
        "SHIP_DF_ART_HOLCAL_10",
        "SHIP_DF_ART_HOLCAL_11",
        "SHIP_DF_ART_HOLCAL_12",
        "SHIP_DF_ART_HOLCAL_2",
        "SHIP_DF_ART_HOLCAL_3",
        "SHIP_DF_ART_HOLCAL_4",
        "SHIP_DF_ART_HOLCAL_5",
        "SHIP_DF_ART_HOLCAL_6",
        "SHIP_DF_ART_HOLCAL_7",
        "SHIP_DF_ART_HOLCAL_8",
        "SHIP_DF_ART_HOLCAL_9",
        "SHIP_DF_ART_MF",
        "SHIP_DF_ART_MTH",
        "SHIP_DF_ART_MTUTH",
        "SHIP_DF_ART_MTUWTH",
        "SHIP_DF_ART_MTUWTHF",
        "SHIP_DF_ART_MTWRSA",
        "SHIP_DF_ART_MW",
        "SHIP_DF_ART_MWF",
        "SHIP_DF_ART_MWTH",
        "SHIP_DF_ART_SUMTUWF",
        "SHIP_DF_ART_SUTH",
        "SHIP_DF_ART_SUTUTH",
        "SHIP_DF_ART_SUTUW",
        "SHIP_DF_ART_SUW",
        "SHIP_DF_ART_TEST",
        "SHIP_DF_ART_TRSA",
        "SHIP_DF_ART_TSA",
        "SHIP_DF_ART_TUF",
        "SHIP_DF_ART_TUFRI",
        "SHIP_DF_ART_TUTH",
        "SHIP_DF_ART_TUTHSA",
        "SHIP_DF_ART_TUWTHF",
        "SHIP_DF_CIGS",
        "SHIP_DF_ECOMM_90",
        "SHIP_DF_FOODS",
        "SHIP_DF_FOODS_805",
        "SHIP_DF_FOODS_86",
        "SHIP_DF_FOODS_STF",
        "SHIP_DF_FRESH_6X_NO_TU_NEXTDAY,MD",
        "SHIP_DF_FRESH_6X_SAMEDAYDMND",
        "SHIP_DF_FRESH_6X_TUE",
        "SHIP_DF_FRESH_6X_WED",
        "SHIP_DF_FRESH_7X",
        "SHIP_DF_FRESH_7X_NEXTDAYDMD",
        "SHIP_DF_FRESH_GAS",
        "SHIP_DF_FRESH_GAS_804",
        "SHIP_DF_FRESH_GAS_872_STRS",
        "SHIP_DF_FRESH_GAS_pilot",
        "SHIP_DF_FRESH_GAS_STORD",
        "SHIP_DF_FRESH_HOL",
        "SHIP_DF_FRESH_MTWF",
        "SHIP_DF_FRESH_MWF",
        "SHIP_DF_FROZEN_860",
        "SHIP_DF_FROZEN_D_ITEMS",
        "SHIP_DF_FROZEN_E_ITEMS",
        "SHIP_DF_FRSH_7X",
        "SHIP_DF_FRSH_C_MTRS",
        "SHIP_DF_FRSH_C_STWF",
        "SHIP_DF_FRZNTEST",
        "SHIP_DF_FRZNTEST 4X",
        "SHIP_DF_FRZN_860",
        "SHIP_DF_FRZN_860 GRP A",
        "SHIP_DF_FRZN_860 GRP B & C",
        "SHIP_DF_FRZN_860_FOR_DC_882",
        "SHIP_DF_FRZN_89_GRP_A 4X",
        "SHIP_DF_FRZN_89_GRP_B 4X",
        "SHIP_DF_FRZN_89_MKT_FMT",
        "SHIP_DF_FRZN_BDG_ST_MKT_FMT",
        "SHIP_DF_FRZN_DEFAULT_MKT_FMT",
        "SHIP_DF_FRZN_GAS",
        "SHIP_DF_FRZN_GAS_804",
        "SHIP_DF_FRZN_GAS_872_STRS",
        "SHIP_DF_FRZN_GAS_STORD",
        "SHIP_DF_FRZN_LANSING_GRP_A",
        "SHIP_DF_FRZN_LANSING_GRP_B",
        "SHIP_DF_FRZN_LANSING_MKT_FMT",
        "SHIP_DF_FRZN_MANU_276",
        "SHIP_DF_FRZN_MANU_872",
        "SHIP_DF_FRZN_NEWSTR_872",
        "SHIP_DF_FRZN_STORD",
        "SHIP_DF_FRZN_STORD_GRP_A",
        "SHIP_DF_FRZN_STORD_GRP_A 4X",
        "SHIP_DF_FRZN_STORD_GRP_B",
        "SHIP_DF_FRZN_STORD_GRP_B 4X",
        "SHIP_DF_FRZN_STORD_NON-LANSING",
        "SHIP_DF_FRZN_STR293",
        "SHIP_DF_GAS_872_SW",
        "SHIP_DF_GM",
        "SHIP_DF_GM_802",
        "SHIP_DF_GM_802_TOYS",
        "SHIP_DF_GM_86_NONART",
        "SHIP_DF_GM_86_STT_NONART",
        "SHIP_DF_GM_90",
        "SHIP_DF_GM_90_MW",
        "SHIP_DF_GM_90_TTH",
        "SHIP_DF_GM_92_802",
        "SHIP_DF_GM_93",
        "SHIP_DF_GM_IN_ARTDMS",
        "SHIP_DF_GROC_STORD_5XA",
        "SHIP_DF_GROC_STORD_5XB",
        "SHIP_DF_GROC_STORD_6X",
        "SHIP_DF_ICECRM_ALLSTR",
        "SHIP_DF_ICECRM_LANSING_STRS",
        "SHIP_DF_ICECRM_NON_LANSING_STR",
        "SHIP_DF_MARKET_MWS",
        "SHIP_DF_NEWSTORE",
        "SHIP_DF_NEWSTORE_10",
        "SHIP_DF_NEWSTORE_11",
        "SHIP_DF_NEWSTORE_12",
        "SHIP_DF_NEWSTORE_13",
        "SHIP_DF_NEWSTORE_14",
        "SHIP_DF_NEWSTORE_15",
        "SHIP_DF_NEWSTORE_16",
        "SHIP_DF_NEWSTORE_17",
        "SHIP_DF_NEWSTORE_18",
        "SHIP_DF_NEWSTORE_19",
        "SHIP_DF_NEWSTORE_2",
        "SHIP_DF_NEWSTORE_20",
        "SHIP_DF_NEWSTORE_21",
        "SHIP_DF_NEWSTORE_22",
        "SHIP_DF_NEWSTORE_23",
        "SHIP_DF_NEWSTORE_24",
        "SHIP_DF_NEWSTORE_25",
        "SHIP_DF_NEWSTORE_26",
        "SHIP_DF_NEWSTORE_27",
        "SHIP_DF_NEWSTORE_28",
        "SHIP_DF_NEWSTORE_29",
        "SHIP_DF_NEWSTORE_3",
        "SHIP_DF_NEWSTORE_30",
        "SHIP_DF_NEWSTORE_4",
        "SHIP_DF_NEWSTORE_5",
        "SHIP_DF_NEWSTORE_6",
        "SHIP_DF_NEWSTORE_7",
        "SHIP_DF_NEWSTORE_8",
        "SHIP_DF_NEWSTORE_86NONART",
        "SHIP_DF_NEWSTORE_9",
        "SHIP_DF_NEWSTORE_ART1",
        "SHIP_DF_NEWSTORE_ART2",
        "SHIP_DF_NEWSTORE_ART3",
        "SHIP_DF_NEWSTORE_ART4",
        "SHIP_DF_NEWSTORE_ART5",
        "SHIP_DF_NEWSTORE_ART6",
        "SHIP_DF_NEWSTORE_FRESH",
        "SHIP_DF_NEWSTORE_FZN",
        "SHIP_DF_NEWSTORE_FZN2",
        "SHIP_DF_PCC_GAS",
        "SHIP_DF_PCC_GAS_SUN_THUR",
        "SHIP_DF_REG_4X2ND",
        "SHIP_DF_REG_4xA_Manu",
        "SHIP_DF_REG_4xA_SO",
        "SHIP_DF_REG_4xB_MANU",
        "SHIP_DF_REG_4xB_SEA",
        "SHIP_DF_REG_4xB_SEA_A",
        "SHIP_DF_REG_4xB_SEA_C",
        "SHIP_DF_REG_4xB_SEA_D",
        "SHIP_DF_REG_4xB_SO",
        "SHIP_DF_REG_4xB_SO_A",
        "SHIP_DF_REG_4xB_SO_C",
        "SHIP_DF_REG_4xB_SO_D",
        "SHIP_DF_REG_5xA_Manu",
        "SHIP_DF_REG_5x_SEA",
        "SHIP_DF_REG_5x_SO",
        "SHIP_DF_REG_SEA_STORD4x",
        "SHIP_DF_REG_STORD_SEA",
        "SHIP_DF_REG_TG",
        "SHIP_DF_RX",
        "SHIP_DF_RXNY",
        "SHIP_DF_RX_CF",
        "SHIP_DF_RX_CF463_RFGR_M",
        "SHIP_DF_RX_CF464_MR",
        "SHIP_DF_RX_CF464_RFGR_T",
        "SHIP_DF_RX_CF_463",
        "SHIP_DF_RX_CF_464",
        "SHIP_DF_RX_NARC",
        "SHIP_DF_RX_NARC_2",
        "SHIP_DF_RX_RFGR_MR",
        "SHIP_DF_RX_RFGR_TF",
        "SHIP_DF_RX_SHIFT",
        "SHIP_DF_SAMEDAY_SUB",
        "SHIP_DF_SAMEDAY_SUB _NO_TU",
        "SHIP_DF_SANDWICH_TRF",
        "SHIP_DF_STRORD_FRSHSEAFD",
        "SHIP_DF_STRORD_FRSHSEAFD_4xStores",
        "SHIP_DF_STRORD_FRSHSEAFD_pilot",
        "SHIP_DF_STRORD_FRSHSEAFD_RGLS",
        "SHIP_DF_STRORD_FRZNSEAFD",
        "SHIP_DF_STRORD_FRZNSEAFD_GRP_A",
        "SHIP_DF_STRORD_FRZNSEAFD_REG_B",
        "SHIP_DF_STRORD_FRZN_PCC",
        "SHIP_DF_SUBSANDWCH",
        "SHIP_DF_SUBSAND_GAS",
        "SHIP_DF_SUBSAND_GAS_pilot",
        "SHIP_DF_SUB_NO_TU",
        "SHIP_DF_SUPPLIES_GAS",
        "SHIP_DF_SUPPLYSTOPFLOW",
        "SHIP_DF_SUPP_460",
        "SHIP_DF_TWRFSA_FIRSTMIN",
    ],
    default=["SHIP_DF_860_TG"]
)

st.markdown(
    """
    ### Additional New Calendars

    Paste one calendar name per line.

    These names will be written directly into the *Cal column
    of the final export file.
    """
)

new_calendar_text = st.text_area(
    "New calendar names",
    height=150,
    placeholder=(
        "SHIP_DF_999_TG\n"
        "SHIP_DF_998_TG\n"
        "SHIP_DF_997_TG"
    )
)

new_calendars = []

for row in new_calendar_text.splitlines():

    cleaned_name = row.strip()

    if cleaned_name:
        new_calendars.append(
            cleaned_name
        )

selected_calendars = (
    existing_calendars
    + new_calendars
)

# Remove accidental duplicates.
selected_calendars = list(
    dict.fromkeys(selected_calendars)
)

if not selected_calendars:

    st.warning(
        "Select or enter at least one calendar."
    )

    st.stop()


# ============================================================
# STEP 3: MINUTE SELECTION
# ============================================================

st.header("Step 3: Select the minutes for each calendar")

minute_selection_by_calendar = {}

for calendar_name in selected_calendars:
    calendar_col, minute_col = st.columns([2, 3])

    with calendar_col:
        st.markdown(f"**{calendar_name}**")

    with minute_col:
        minute_selection_by_calendar[calendar_name] = st.selectbox(
            f"Minutes for {calendar_name}",
            options=MINUTE_OPTIONS,
            key=f"minute_option__{safe_key(calendar_name)}",
            label_visibility="collapsed",
        )


# ============================================================
# STEP 4: CLICKABLE CALENDARS
# ============================================================

st.header("Step 4: Mark calendar dates")

st.info(
    "Each date starts gray and is excluded from export. Click once to mark it "
    "green/Open, click again to mark it red/Closed, and click a third time to "
    "return it to gray/Not selected. Include seven normally scheduled days at "
    "the end when a 7D automatic repeat should be added."
)

st.caption("Click a date to cycle: ⚪ Not Selected → 🟢 Open → 🔴 Closed")

for calendar_name in selected_calendars:
    st.subheader(calendar_name)
    st.caption(
        f"Minute rule: {minute_selection_by_calendar[calendar_name]}"
    )

    action_col1, action_col2, action_col3, spacer = st.columns([1, 1, 1, 3])

    with action_col1:
        if st.button(
            "Mark all Open",
            key=f"all_open__{safe_key(calendar_name)}",
            use_container_width=True,
        ):
            mark_all_calendar_dates(calendar_name, fiscal_weeks, "Open")
            st.rerun()

    with action_col2:
        if st.button(
            "Mark all Closed",
            key=f"all_closed__{safe_key(calendar_name)}",
            use_container_width=True,
        ):
            mark_all_calendar_dates(calendar_name, fiscal_weeks, "Closed")
            st.rerun()

    with action_col3:
        if st.button(
            "Clear all",
            key=f"clear_all__{safe_key(calendar_name)}",
            use_container_width=True,
        ):
            clear_calendar_dates(calendar_name, fiscal_weeks)
            st.rerun()

    header_columns = st.columns([0.7, 1, 1, 1, 1, 1, 1, 1])
    header_columns[0].markdown("**Fiscal Week**")

    for index, weekday_name in enumerate(WEEKDAY_NAMES, start=1):
        header_columns[index].markdown(
            f"<div style='text-align:center'><b>{weekday_name}</b></div>",
            unsafe_allow_html=True,
        )

    for week_number in fiscal_weeks:
        row_columns = st.columns([0.7, 1, 1, 1, 1, 1, 1, 1])
        row_columns[0].markdown(f"### FW {week_number}")

        week_dates = dates_for_fiscal_week(week_number)

        for day_index, calendar_date in enumerate(week_dates, start=1):
            status = get_day_status(calendar_name, calendar_date)
            style = STATUS_STYLE[status]
            button_label = (
                f"{style['icon']} {calendar_date.strftime('%b %d')}   "
                f"{style['label']}"
            )


            with row_columns[day_index]:
                if st.button(
                    button_label,
                    key=(
                        f"day_button__{safe_key(calendar_name)}__"
                        f"{calendar_date.isoformat()}"
                    ),
                    use_container_width=True,
                   
                ):
                    cycle_day_status(calendar_name, calendar_date)
                    st.rerun()

    counts = selected_date_summary(calendar_name, fiscal_weeks)
    summary_col1, summary_col2, summary_col3 = st.columns(3)
    summary_col1.metric("Open dates", counts["Open"])
    summary_col2.metric("Closed dates", counts["Closed"])
    summary_col3.metric("Not selected", counts["Not selected"])

    st.divider()


# ============================================================
# STEP 5: DESCRIPTION, PREVIEW, AND DOWNLOAD
# ============================================================

st.header("Step 5: Describe and download the file")

description = st.text_input(
    "Description for the Descr column",
    value="turkeys 2026",
)

add_repeat = st.checkbox(
    "Add the Opt 5 / 7D repeat row for each calendar",
    value=True,
    help=(
        "Adds one new row for the day immediately after the "
        "latest selected date. The row receives Opt = 5 and "
        "Repeat = 7D."
    ),
)

output_df = build_output_dataframe(
    selected_calendars=selected_calendars,
    fiscal_weeks=fiscal_weeks,
    minute_selection_by_calendar=minute_selection_by_calendar,
    description=description,
    add_repeat=add_repeat,
)

if output_df.empty:
    st.warning(
        "No dates are currently marked Open or Closed. Gray dates are not exported."
    )
else:
    preview_df = output_df.copy()
    preview_df["*Eff"] = preview_df["*Eff"].apply(
        lambda value: (
            f"{value.month}/{value.day}/{value.year} "
            f"{value.hour}:{value.minute:02d}"
        )
    )

    st.subheader("Output preview")
    st.dataframe(
        preview_df,
        hide_index=True,
        use_container_width=True,
        height=450,
    )

    st.caption(f"The output contains {len(output_df):,} records.")

    output_name = st.text_input(
        "Output filename",
        value="Calendar_Import",
    )

    safe_output_name = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        output_name.strip(),
    ).strip("_")

    if not safe_output_name:
        safe_output_name = "Calendar_Import"

    if not description.strip():
        st.error("Enter a description before downloading.")
    else:
        excel_file = create_excel_file(output_df)

        st.download_button(
            label="📥 Download Calendar Import File",
            data=excel_file,
            file_name=f"{safe_output_name}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
