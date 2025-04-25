import streamlit as st
import pandas as pd
import pdfplumber

def pdf_to_excel(pdf_file, output_file):
    try:
        # Open the PDF file
        pages = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                pages.append(text)
        
        # Create a pandas DataFrame from the extracted text
        df = pd.DataFrame(pages)
        
        # Write the DataFrame to an Excel file
        df.to_excel(output_file, index=False)
    except FileNotFoundError:
        st.error("Error: PDF file not found")
    except Exception as e:
        st.error(f"An error occurred: {e}")

st.title("PDF to Excel Converter")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    pdf_file = uploaded_file.name
    output_file = "output.xlsx"
    
    st.write("Converting PDF to Excel...")
    pdf_to_excel(pdf_file, output_file)
    
    st.download_button(label="Download Excel file", data=st.as_folium(output_file), file_name='output.xlsx')

if not uploaded_file:
    st.write("No file selected")