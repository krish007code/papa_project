"""
Streamlit front-end for the OLT Down Report Generator.
Run with:  streamlit run app.py
"""

import io
from datetime import datetime

import streamlit as st

from oltreport import generate_report

st.set_page_config(
    page_title="OLT Down Report Generator",
    page_icon="📡",
    layout="centered",
)

st.title("OLT Down Report Generator")
st.caption("Upload the three required Excel files and click **Generate Report**.")

st.divider()

col1, col2 = st.columns(2)

with col1:
    permanent_file = st.file_uploader(
        "🗂️ Permanent file",
        type="xlsx",
        help="BA_OA_UP_WEST_LIST.xlsx — never changes",
    )
    monthly_file = st.file_uploader(
        "📅 Monthly file",
        type="xlsx",
        help="AMC_GP_Status_June_26_*.xlsx — updated monthly",
    )

with col2:
    daily_file = st.file_uploader(
        "📋 Daily report file",
        type="xlsx",
        help="report_<id>_<date>.xlsx — downloaded every day",
    )
    generated_by = st.text_input("👤 Username", value="bsupwag2")

st.divider()

all_uploaded = permanent_file and monthly_file and daily_file

if st.button("⚙️ Generate Report", disabled=not all_uploaded, type="primary"):
    with st.spinner("Processing…"):
        try:
            excel_bytes = generate_report(
                permanent_path=io.BytesIO(permanent_file.read()),
                monthly_path=io.BytesIO(monthly_file.read()),
                daily_path=io.BytesIO(daily_file.read()),
                now=datetime.now(),
                generated_by=generated_by,
            )
            today_fn = datetime.now().strftime("%Y-%m-%d")
            st.success("✅ Report generated successfully!")
            st.download_button(
                label="⬇️ Download OLT_Down_Report.xlsx",
                data=excel_bytes,
                file_name=f"OLT_Down_Report_{today_fn}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except FileNotFoundError as e:
            st.error(f"File error: {e}")
        except ValueError as e:
            st.error(f"Data error: {e}")
        except Exception as e:
            st.error(f"Unexpected error: {e}")
            raise

if not all_uploaded:
    st.info("⬆️ Upload all three files above to enable the Generate button.")