import streamlit as st
import pandas as pd

st.title("3 Worst Drawdowns")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        st.write("Preview of the uploaded data:")
        st.dataframe(df)

        column = st.selectbox("Select a column to analyze", df.columns)

        if column:
            try:
                numeric_col = pd.to_numeric(df[column], errors='coerce')
                top_values = numeric_col.dropna().nlargest(3)

                st.write(f"Top 3 highest values in '{column}':")
                st.write(top_values)
            except Exception as e:
                st.error(f"Error processing column: {e}")
    except Exception as e:
        st.error(f"Error reading the file: {e}")