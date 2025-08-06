import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Merge Excel Files into One", layout="centered")

st.title("📊 Merge Multiple Excel Files into One (with Multiple Sheets)")
st.markdown("Upload your Excel files. Each will be stored in its own sheet in a single file.")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for file in uploaded_files:
            try:
                df = pd.read_excel(file)
                sheet_name = file.name[:31]  # Excel sheet name max length = 31
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            except Exception as e:
                st.error(f"❌ Error processing file: {file.name} — {str(e)}")
    
    st.success("✅ Files merged successfully!")
    
    st.download_button(
        label="📥 Download Merged Excel File",
        data=output.getvalue(),
        file_name="merged_commodity_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
