import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from io import BytesIO

# -------------------------------------------------
# Page setup
# -------------------------------------------------
st.set_page_config(page_title="Rota Generator", layout="wide")
st.title("📅 Rota Generator")
st.write("Upload the completed rota template and download an Excel rota.")

# -------------------------------------------------
# Upload
# -------------------------------------------------
uploaded = st.file_uploader("Upload rota_generator_template.xlsx", type=["xlsx"])
if not uploaded:
    st.stop()

# -------------------------------------------------
# Load & validate sheets
# -------------------------------------------------
xls = pd.ExcelFile(uploaded)
REQUIRED_SHEETS = ["Staff", "WorkingHours", "Holidays", "Parameters"]
missing = [s for s in REQUIRED_SHEETS if s not in xls.sheet_names]

if missing:
    st.error(f"Missing required sheets: {', '.join(missing)}")
    st.stop()

staff_df = pd.read_excel(xls, "Staff")
hours_df = pd.read_excel(xls, "WorkingHours")
holidays_df = pd.read_excel(xls, "Holidays")
params_df = pd.read_excel(xls, "Parameters")

st.success("Template loaded")

# -------------------------------------------------
# Helpers
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
# Parameters
# -------------------------------------------------
params = dict(zip(params_df["Rule"], params_df["Value"]))
BREAK_START = parse_time(params.get("Break_Window_Start", "11:30"))
BREAK_END = parse_time(params.get("Break_Window_End", "14:00"))

# -------------------------------------------------
# Staff
# -------------------------------------------------
staff_df["Staff_ID"] = staff_df["Staff_ID"].astype(str)

STAFF = {}
for _, r in staff_df.iterrows():
    STAFF[r["Staff_ID"]] = {
        "name": r["Name"],
        "site": r["HomeSite"],
        "skills": {
            "front_desk": r.get("FrontDesk", "N") == "Y",
            "admin": r.get("Triage", "N") == "Y",
        },
    }

# -------------------------------------------------
# Working hours
# -------------------------------------------------
WORKING = {}
for _, r in hours_df.iterrows():
    sid = str(r["Staff_ID"])
    WORKING.setdefault(r["Day"], {})[sid] = (
        parse_time(r["Start_Time"]),
        parse_time(r["End_Time"]),
    )

# -------------------------------------------------
# Holidays
# -------------------------------------------------
HOLIDAYS = {}
for _, r in holidays_df.iterrows():
    HOLIDAYS.setdefault(str(r["Staff_ID"]), []).append(
        (
            pd.to_datetime(r["StartDate"]).date(),
            pd.to_datetime(r["EndDate"]).date(),
        )
    )


def on_holiday(sid, d):
    for s, e in HOLIDAYS.get(sid, []):
        if s <= d <= e:
            return True
    return False


# -------------------------------------------------
# Rota generation (FIXED FUNCTION)
# -------------------------------------------------
def generate_day(date):
    weekday = date.strftime("%A")
    blocks = build_blocks()
    rota = {b: [] for b in blocks}

    available = []
    for sid, (s, e) in WORKING.get(weekday, {}).items():
        if not on_holiday(sid, date):
            available.append((sid, s, e))

    for block in blocks:
        for site in ["SLGP", "JEN", "BGS"]:
            for sid, s, e in available:
                staff = STAFF[sid]
                if (
                    staff["site"] == site
                    and staff["skills"]["front_desk"]
                    and in_range(block, s, e)
                ):
                    rota[block].append((sid, f"Front_Desk_{site}"))
                    break

        if block < time(16, 0):
            for site in ["SLGP", "JEN"]:
                for sid, s, e in available:
                    staff = STAFF[sid]
                    if (
                        staff["site"] == site
                        and staff["skills"]["admin"]
                        and in_range(block, s, e)
                    ):
                        rota[block].append((sid, f"Triage_Admin_{site}"))
                        break

    return rota


# -------------------------------------------------
# Generate + Export
# -------------------------------------------------
if st.button("Generate rota"):
    start_date = pd.to_datetime(params["Week_Commencing"]).date()
    dates = [start_date + timedelta(days=i) for i in range(5)]

    wb = Workbook()
    wb.remove(wb.active)

    for d in dates:
        ws = wb.create_sheet(d.strftime("%A"))
        ws.append(["Time", "Assignments"])

        rota = generate_day(d)
        for block in sorted(rota):
            ws.append(
                [
                    block.strftime("%H:%M"),
                    ", ".join(f"{STAFF[s]['name']} ({r})" for s, r in rota[block]),
                ]
            )

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "📊 Download rota.xlsx",
        data=output,
        file_name="rota.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
