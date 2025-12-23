import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from io import BytesIO

st.set_page_config(page_title="Rota Generator", layout="wide")

st.title("📅 Rota Generator")
st.write("Upload your completed template, generate a rota, and download it as Excel.")

uploaded = st.file_uploader(
    "Upload rota_generator_template.xlsx",
    type=["xlsx"]
)

if uploaded:
    staff = pd.read_excel(uploaded, sheet_name="Staff")
    hours = pd.read_excel(uploaded, sheet_name="WorkingHours")
    holidays = pd.read_excel(uploaded, sheet_name="Holidays")
    skills = pd.read_excel(uploaded, sheet_name="Skills")
    params = pd.read_excel(uploaded, sheet_name="Parameters")

    st.success("Template loaded successfully")

    if st.button("Generate rota"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Rota_Grid"

        ws.append(["Time", "Example"])
        t = datetime.strptime("08:00", "%H:%M")
        end = datetime.strptime("18:30", "%H:%M")

        while t < end:
            ws.append([t.strftime("%H:%M"), "Generated"])
            t += timedelta(minutes=30)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            "📊 Download rota.xlsx",
            data=output,
            file_name="rota.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
