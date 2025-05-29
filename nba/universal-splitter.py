import streamlit as st
import pandas as pd
import numpy as np
import io
import random

st.set_page_config(layout="wide")

st.sidebar.title("🔧 Single Column Marks Splitter")

# Sidebar Inputs
divisions = st.sidebar.number_input("Number of divisions", min_value=1, value=4)
division_type = st.sidebar.selectbox("Division type", ["Equal", "Random"])
max_component = st.sidebar.number_input("Max marks per component (0 = no limit)", min_value=0.0, value=0.0)
uploaded_file = st.sidebar.file_uploader("Upload Excel file (with 'marks' column)", type=["xlsx"])
process_btn = st.sidebar.button("🚀 Process")

st.title("📊 Marks Division App")
st.write("This tool takes a single column of marks and divides each value into multiple components.")

# Function to divide marks
def split_marks(total, divisions, division_type, max_component):
    total = round(float(total), 2)
    is_decimal = not float(total).is_integer()

    if division_type == "Equal":
        base = round(total / divisions, 2)
        result = [base] * divisions
        diff = round(total - sum(result), 2)
        result[-1] = round(result[-1] + diff, 2)

        if not is_decimal:
            result = [int(x) for x in result]

        if max_component > 0:
            result = [min(x, max_component) for x in result]
            current_sum = round(sum(result), 2)
            if current_sum < total:
                remaining = round(total - current_sum, 2)
                for i in range(len(result)):
                    space = max_component - result[i]
                    addable = min(space, remaining)
                    result[i] = round(result[i] + addable, 2)
                    remaining = round(remaining - addable, 2)
                    if remaining <= 0:
                        break
            if not is_decimal:
                result = [int(x) for x in result]

    else:  # Random division
        if total == 0:
            return [0] * divisions

        int_part = int(total)
        frac_part = round(total - int_part, 2)

        if int_part < divisions:
            parts = [1]*int_part + [0]*(divisions - int_part)
            random.shuffle(parts)
        else:
            points = sorted(random.sample(range(1, int_part + divisions), divisions - 1))
            parts = [points[0]]
            for i in range(1, len(points)):
                parts.append(points[i] - points[i-1])
            parts.append(int_part + divisions - points[-1])
            parts = [p - 1 for p in parts]

        if is_decimal:
            idx = random.randint(0, divisions - 1)
            parts[idx] = round(parts[idx] + frac_part, 2)

        result = parts
        result = [round(max(0, val), 2) for val in result]

        diff = round(total - sum(result), 2)
        if abs(diff) > 0.01:
            result[-1] = round(result[-1] + diff, 2)

        if not is_decimal:
            result = [int(round(x)) for x in result]

        # Apply max per component
        if max_component > 0:
            result = [min(x, max_component) for x in result]
            current_sum = round(sum(result), 2)
            if current_sum < total:
                remaining = round(total - current_sum, 2)
                for i in range(len(result)):
                    space = max_component - result[i]
                    addable = min(space, remaining)
                    result[i] = round(result[i] + addable, 2)
                    remaining = round(remaining - addable, 2)
                    if remaining <= 0:
                        break
            if not is_decimal:
                result = [int(x) for x in result]

    return result

# Main processing
if process_btn and uploaded_file:
    df = pd.read_excel(uploaded_file)
    if 'marks' not in df.columns:
        st.error("Uploaded file must contain a 'marks' column.")
    else:
        result_df = df.copy()
        for i in range(divisions):
            result_df[f'q{i+1}'] = df['marks'].apply(lambda x: split_marks(x, divisions, division_type, max_component)[i])

        st.success("✅ Processing complete!")
        st.dataframe(result_df)

        # Download
        output = io.BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("📥 Download Result File", data=output,
                           file_name="divided_marks.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
