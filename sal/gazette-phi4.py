import streamlit as st
import pdfplumber
import pandas as pd

def extract_tables_from_pdf(pdf_path):
    tables = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract all tables on the current page
            page_tables = page.extract_tables()
            if page_tables:
                # Append each table to the list of tables
                tables.extend(page_tables)
                
    return tables

def convert_to_dataframe(tables):
    # Convert the list of tables (list of lists) into a single DataFrame
    all_rows = []
    
    for table in tables:
        df = pd.DataFrame(table[1:], columns=table[0])  # Skip header row in each table
        all_rows.append(df)
        
    if all_rows:
        return pd.concat(all_rows, ignore_index=True)
    else:
        return None

def main():
    st.title("PDF to Excel Converter with pdfplumber")

    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    if uploaded_file is not None:
        pdf_path = 'temp.pdf'
        
        # Save the uploaded file temporarily
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        tables = extract_tables_from_pdf(pdf_path)
        
        if tables:
            df = convert_to_dataframe(tables)
            
            st.success("PDF successfully converted to DataFrame!")
            st.dataframe(df)

            # Convert DataFrame to Excel
            excel_file_path = 'output.xlsx'
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            with open(excel_file_path, "rb") as file:
                st.download_button(
                    label="Download Excel",
                    data=file,
                    file_name=excel_file_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("No extractable tables found in the PDF.")

if __name__ == "__main__":
    main()