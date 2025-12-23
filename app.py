import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from io import BytesIO

# -------------------------------------------------
# Page setup
# -------------------------------------------------
st.set_page_config(
    page_title="Rota Generator",
    layout="wide"
)

st.title("📅 Rota Generator")
st.write(
    "Upload the completed rota template, generate the rota, "
    "and download the Excel output."
)

# -------------------------------------------------
# Upload
# -------------------------------------------------
uploaded = st.file_uploader(
    "Upload rota_generator_template.xlsx",
    type=["xlsx"]
)

if not uploaded:
    st.info("Please upload the completed template to continue.")
    st.stop()

# -------------------------------------------------
# Load & validate sheets
# -------------------------------------------------
xls = pd.ExcelFile(uploaded)

REQUIRED_SHEETS = [
    "Staff",
    "WorkingHours",
    "Holidays",
    "Parameters",
]

missing = [s for s in REQUIRED_SHEETS if s not in xls.sheet_names]
if missing:
    st.error(f"Missing required sheets: {', '.join(missing)}")
    st.stop()

staff_df = pd.read_excel(xls, "Staff")
hours_df = pd.read_excel(xls, "WorkingHours")
holidays_df = pd.read_excel(xls, "Holidays")
params_df = pd.read_excel(xls, "Parameters")

st.success("Template loaded successfully")

# -------------------------------------------------
# Helper functions
# -------------------------------------------------
def parse_time(val):
    if pd.isna(val):
        return None
    if isinstance(val, time):
        return val
    return datetime.strptime(str(val), "%H:%M").time()


def build_blocks():
    blocks = []
    current = datetime.combine(datetime.today(), time(8, 0))
    end = datetime.combine(datetime.today(), time(18, 30))
    while current < end:
        blocks.append(current.time())
        current += timedelta(minutes=30)
    return blocks


def in_range(t, start, end):
    return start <= t < end


# -------------------------------------------------
# Extract parameters
# -------------------------------------------------
params = dict(zip(params_df["Rule"], params_df["Value"]))

BREAK_START = parse_time(params.get("Break_Window_Start", "11:30"))
BREAK_END = parse_time(params.get("Break_Window_End", "14:00"))

# -------------------------------------------------
# Prepare staff data
# -------------------------------------------------
staff_df["Staff_ID"] = staff_df["Staff_ID"].astype(str)

STAFF = {}
for _, row in staff_df.iterrows():
    STAFF[row["Staff_ID"]] = {
        "name": row["Name"],
        "site": row["HomeSite"],
        "skills": {
            "front_desk": row.get("FrontDesk", "N") == "Y",
            "admin": row.get("Triage", "N") == "Y",
            "email": row.get("Email", "N") == "Y",
            "phones": row.get("Phones", "N") == "Y",
            "bookings": row.get("Bookings", "N") == "Y",
            "emis": row.get("EMIS", "N") == "Y",
            "docman": row.get("Docman", "N") == "Y",
        },
        "is_carol": row.get("IsCarolChurch", "N") == "Y",
    }

# -------------------------------------------------
# Working hours by weekday
# -------------------------------------------------
WORKING = {}
for _, row in hours_df.iterrows():
    sid = str(row["Staff_ID"])
    day = row["Day"]
    start = parse_time(row["Start_Time"])
    end = parse_time(row["End_Time"])
    WORKING.setdefault(day, {})[sid] = (start, end)

# -------------------------------------------------
# Holidays
# -------------------------------------------------
HOLIDAYS = {}
for _, row in holidays_df.iterrows():
    sid = str(row["Staff_ID"])
    HOLIDAYS.setdefault(sid, []).append(
        (
            pd.to_datetime(row["StartDate"]).date(),
            pd.to_datetime(row["EndDate"]).date(),
        )
    )


def on_holiday(staff_id, date):
    for start, end in HOLIDAYS.get(staff_id, []):
        if start <= date <= end:
            return True
    return False


# -------------------------------------------------
# Rota generation (safe, minimal, extendable)
# -------------------------------------------------
def gene
