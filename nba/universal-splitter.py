import streamlit as st
import pandas as pd
import io
import random
from math import floor

st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

# Sample input for download
sample_data = pd.DataFrame({'marks': [39.5, 38, 36.5, 40, 21]})
sample_buffer = io.BytesIO()
sample_data.to_excel(sample_buffer, index=False)
sample_buffer.seek(0)
st.sidebar.download_button(
    label="📥 Download Sample Input File",
    data=sample_buffer,
    file_name="sample_input.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_files = st.sidebar.file_uploader(
    "Upload Excel files (1 column: marks)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

num_divisions = st.sidebar.number_input(
    "Number of divisions", min_value=1, max_value=100, value=5, step=1
)
division_type = st.sidebar.selectbox("Division type", ["Equal", "Random"])
max_per_component = st.sidebar.number_input(
    "Max marks per component (optional, 0 disables)", 
    min_value=0.0, value=10.0, step=0.25,
    help="Max marks per division component. 0 disables limit."
)

process_button = st.sidebar.button("Start Processing")

st.title("📊 Marks Division Panel")
st.header(":green[Divide 'marks' Column into Divisions]")
st.subheader(":blue[Equal and Random Divisions with neat decimals (0.25 steps)]")

def round_to_step(value, step=0.25):
    return round(value / step) * step

def split_marks_with_caps_step(total, divisions, mode, max_cap, step=0.25):
    total = round(float(total), 2)
    caps = [max_cap]*divisions if max_cap > 0 else [total]*divisions
    caps = [round(c, 2) for c in caps]
    result = [0.0] * divisions

    if total == 0:
        return result

    # Fix sum helper
    def fix_sum(arr, target):
        diff = round(target - sum(arr), 2)
        i = 0
        while abs(diff) >= step and i < len(arr):
            if diff > 0:
                space = caps[i] - arr[i]
                if space >= step:
                    change = min(space, diff)
                    change = floor(change / step) * step
                    if change <= 0:
                        i += 1
                        continue
                    arr[i] += change
            else:
                if arr[i] >= step:
                    change = min(arr[i], abs(diff))
                    change = floor(change / step) * step
                    if change <= 0:
                        i += 1
                        continue
                    arr[i] -= change
            arr[i] = round(arr[i], 2)
            diff = round(target - sum(arr), 2)
            i += 1
        return arr

    if mode == "Equal":
        base = total / divisions
        for i in range(divisions):
            val = min(base, caps[i])
            result[i] = round_to_step(val, step)

        result = fix_sum(result, total)

    else:  # Random
        remaining = total
        attempts = 0
        while remaining >= step and attempts < 10000:
            i = random.randint(0, divisions - 1)
            space = caps[i] - result[i]
            if space < step:
                attempts += 1
                continue
            add = round_to_step(random.uniform(step, min(space, remaining)), step)
            if add < step:
                attempts += 1
                continue
            result[i] += add
            result[i] = round(result[i], 2)
            remaining = round(remaining - add, 2)
            attempts += 1

        result = fix_sum(result, total)

    # Final cap check & non-negative
    result = [min(round(val, 2), caps[i]) for i, val in enumerate(result)]
    result = [max(0, v) for v in result]

    return result

def process_row(row):
    out = {}
    val = row.get('marks')
    invalid = False
    try:
        val = float(val)
        if val < 0:
            invalid = True
    except:
        invalid = True
        val = 0

    # Check if original marks have decimal part
    has_decimal = (val != int(val))

    # Use step 0.25 if decimals exist, else 1
    step = 0.25 if has_decimal else 1.0

    # If max_per_component=0, disable cap by setting to large number
    cap = max_per_component if max_per_component > 0 else val + 1000

    divisions_values = split_marks_with_caps_step(val, num_divisions, division_type, cap, step)

    # If original marks are integer, convert all divisions to int
    if not has_decimal:
        divisions_values = [int(round(v)) for v in divisions_values]

    for i, v in enumerate(divisions_values, start=1):
        out[f"div_{i}"] = v

    out['marks'] = val
    return out, invalid

def process_file(file):
    df = pd.read_excel(file)
    if 'marks' not in df.columns:
        st.warning(f"File '{file.name}' does not contain 'marks' column.")
        return None
    processed = []
    invalid_rows = []

    for idx, row in df.iterrows():
        processed_row, is_invalid = process_row(row)
        processed.append(processed_row)
        if is_invalid:
            invalid_rows.append(idx)

    out_df = pd.DataFrame(processed)

    buffer = io.BytesIO()
    out_df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer, invalid_rows, file.name

if process_button and uploaded_files:
    for file in uploaded_files:
        processed_buffer, invalid_rows, filename = process_file(file)
        if processed_buffer is not None:
            st.success(f"✅ Processed: {filename}")
            st.download_button(
                label=f"📥 Download processed '{filename}'",
                data=processed_buffer,
                file_name=f"processed_{filename}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=filename
            )
            if invalid_rows:
                st.warning(f"⚠️ Some invalid rows detected in '{filename}': {invalid_rows}")

if not uploaded_files:
    st.info("Upload one or more Excel files with a single column named 'marks' to start.")
