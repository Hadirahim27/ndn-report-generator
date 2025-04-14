
import streamlit as st
from ndn_report_generator import process_ndn_excel

st.set_page_config(page_title="NDN Report Generator", layout="centered")
st.title("📊 NDN Report Generator")
st.write("Upload the raw Excel file to generate the cleaned report.")

uploaded_file = st.file_uploader("📂 Upload NDN Raw Excel File", type=["xlsx"])

if uploaded_file:
    with st.spinner("⏳ Processing data..."):
        result = process_ndn_excel(uploaded_file)
    st.success("✅ Report generated successfully!")

    st.download_button(
        label="📥 Download Cleaned Report",
        data=result,
        file_name="NDN_Cleaned_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Please upload a valid .xlsx file to begin.")
