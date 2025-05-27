import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Uremove", layout="wide")
st.title("Un-attempted Question Removal")
st.header(":green[ST1, ST2, ETE]", divider="rainbow")
st.subheader(":red[Removes the un-attempted questions from ST1, ST2, ETE]", divider="rainbow")
st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

uploaded_file = st.sidebar.file_uploader("Upload the Excel File", type=["xlsx"])

def clean_and_shift(df, columns_out, columns_in):
    df_copy = df.copy()
    input_data = df_copy[columns_in].copy()

    for i in df.index:
        values = [v for v in input_data.loc[i] if v != 'U']
        values = values[:len(columns_out)]
        values += [''] * (len(columns_out) - len(values))
        df_copy.loc[i, columns_out] = values

    return df_copy

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df = df.replace("N/A", "U").fillna("U")

    # --- ST1 ---
    df = clean_and_shift(df,
                         ['st1-2', 'st1-3', 'st1-4', 'st1-5', 'st1-6'],
                         ['st1-2', 'st1-3', 'st1-4', 'st1-5', 'st1-6', 'st1-7'])

    df = clean_and_shift(df,
                         ['st1-7', 'st1-8', 'st1-9'],
                         ['st1-8', 'st1-9', 'st1-10', 'st1-11'])

    df = clean_and_shift(df,
                         ['st1-10'],
                         ['st1-12', 'st1-13'])

    # --- ST2 ---
    df = clean_and_shift(df,
                         ['st2-2', 'st2-3', 'st2-4', 'st2-5', 'st2-6'],
                         ['st2-2', 'st2-3', 'st2-4', 'st2-5', 'st2-6', 'st2-7'])

    df = clean_and_shift(df,
                         ['st2-7', 'st2-8', 'st2-9'],
                         ['st2-8', 'st2-9', 'st2-10', 'st2-11'])

    df = clean_and_shift(df,
                         ['st2-10'],
                         ['st2-12', 'st2-13'])

    # --- ETE ---
    df = clean_and_shift(df,
                         ['ete-q2', 'ete-q3', 'ete-q4', 'ete-q5', 'ete-q6'],
                         ['ete-q2', 'ete-q3', 'ete-q4', 'ete-q5', 'ete-q6', 'ete-q7'])

    df = clean_and_shift(df,
                         ['ete-q7', 'ete-q8', 'ete-q9', 'ete-q10', 'ete-q11'],
                         ['ete-q8', 'ete-q9', 'ete-q10', 'ete-q11', 'ete-q12', 'ete-q13'])

    df = clean_and_shift(df,
                         ['ete-q12', 'ete-q13'],
                         ['ete-q14', 'ete-q15', 'ete-q16'])

    # ❌ Remove unneeded columns
    columns_to_drop = [
        'st1-11', 'st1-12', 'st1-13',
        'st2-11', 'st2-12', 'st2-13',
        'ete-q14', 'ete-q15', 'ete-q16'
    ]
    df = df.drop(columns=[col for col in columns_to_drop if col in df.columns])

    # ✅ Downloadable file
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("✅ All processing complete and unused columns removed.")

    st.sidebar.download_button(
        "📥 Download Cleaned Excel",
        output,
        "Cleaned_Marks_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
