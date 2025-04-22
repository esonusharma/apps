import streamlit as st
import pandas as pd
import pdfplumber
import io

def extract_tables_from_pdf(pdf):
    """
    Extracts tables from a PDF object using pdfplumber.
    Handles potential errors and builds the DataFrame manually.
    """
    try:
        all_data = []
        for page in pdf.pages:
            for table in page.extract_tables():
                if table:  # Ensuring there's actually a table to add
                    all_data.extend(table)

        if not all_data:
            st.warning("No tables found in the PDF. The PDF might not contain any tables or the tables might be heavily formatted.")
            return None

        df = pd.DataFrame(all_data)
        return df

    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None


def convert_df_to_excel(df):
    """
    Converts a DataFrame to an Excel file in memory.
    """
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Extracted Data", index=False)
    excel_buffer.seek(0)
    return excel_buffer


# Streamlit App
st.title("PDF Table Extractor")

uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"])

if uploaded_file is not None:
    # Convert the uploaded file to a bytes object and open it with pdfplumber
    pdf_bytes = io.BytesIO(uploaded_file.read())
    
    with pdfplumber.open(pdf_bytes) as pdf:
        df = extract_tables_from_pdf(pdf)
        
        if df is not None:
            st.dataframe(df)

            # Add a download button for the Excel file
            excel_buffer = convert_df_to_excel(df)
            st.download_button(
                label="Download data as Excel",
                data=excel_buffer,
                file_name='extracted_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )