import streamlit as st
import pandas as pd
import numpy as np
import io
import random

st.title("Random Marks Splitter (13 Columns)")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

def distribute_marks(total):
    # Define column structure
    structure = {
        1: 5,
        2: 2, 3: 2, 4: 2, 5: 2, 6: 2, 7: 2,
        8: 5, 9: 5, 10: 5, 11: 5,
        12: 10, 13: 10
    }

    # Select one column to be N/A from each group
    na_2_7 = random.choice(range(2, 8))
    na_8_11 = random.choice(range(8, 12))
    na_12_13 = random.choice(range(12, 14))
    na_columns = {na_2_7, na_8_11, na_12_13}

    # Filter only columns that are not N/A
    valid_columns = [col for col in range(1, 14) if col not in na_columns]
    max_values = [structure[col] for col in valid_columns]

    # Random distribution within max values such that sum = total
    while True:
        values = [random.uniform(0, mx) for mx in max_values]
        scale = total / sum(values) if sum(values) != 0 else 0
        scaled = [round(v * scale, 1) for v in values]
        if sum(scaled) == total:
            break

    result = {}
    index = 0
    for col in range(1, 14):
        if col in na_columns:
            result[str(col)] = 'N/A'
        else:
            result[str(col)] = scaled[index]
            index += 1

    return result

if uploaded_files:
    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)
        if 'marks' not in df.columns:
            st.error(f"'marks' column not found in {uploaded_file.name}")
            continue

        output_data = df.copy()
        split_columns = []

        for idx, row in df.iterrows():
            split = distribute_marks(row['marks'])
            split_columns.append(split)

        split_df = pd.DataFrame(split_columns)
        final_df = pd.concat([df, split_df], axis=1)

        output = io.BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label=f"Download Processed File: {uploaded_file.name}",
            data=output,
            file_name=f"processed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
