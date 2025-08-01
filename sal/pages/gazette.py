import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="PDF Table to Excel", layout="wide")
st.title("📊 Extract Tables from High-Quality PDFs to Excel")

uploaded_file = st.file_uploader("Upload PDF file with tables", type=["pdf"])

if uploaded_file:
    tables = []
    st.info("🔍 Scanning PDF for tables...")

    with pdfplumber.open(uploaded_file) as pdf:
        for i, page in enumerate(pdf.pages):
            st.write(f"📄 Page {i + 1}")
            page_tables = page.extract_tables()
            if not page_tables:
                st.warning("No tables found on this page.")
            for table in page_tables:
                df = pd.DataFrame(table[1:], columns=table[0])
                if not df.empty:
                    st.dataframe(df)
                    tables.append(df)

    if tables:
        final_df = pd.concat(tables, ignore_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Extracted Tables')

        st.success("✅ Tables extracted successfully!")
        st.download_button(
            label="📥 Download as Excel",
            data=output.getvalue(),
            file_name="tables_from_pdf.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ No tables found in the uploaded PDF.")
