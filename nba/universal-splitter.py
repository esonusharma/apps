import streamlit as st
import pandas as pd
import io
import random

st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

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

def round_step(x, step=0.25):
    """Round x to nearest multiple of step"""
    return round(x / step) * step

def distribute_equal(total, divisions, max_per_comp, step=0.25):
    # If max_per_comp == 0, no cap
    if max_per_comp <= 0 or max_per_comp >= total:
        # Just divide equally and round
        base = total / divisions
        res = [round_step(base, step)] * divisions
    else:
        # Cap at max_per_comp per division
        base = min(total / divisions, max_per_comp)
        res = [round_step(base, step)] * divisions

    # Adjust sum to total by distributing leftover carefully
    current_sum = sum(res)
    diff = round(total - current_sum, 2)

    # Increase divisions one by one if diff > 0 and can add without exceeding cap
    i = 0
    while diff >= step and i < divisions:
        addable = max_per_comp - res[i] if max_per_comp > 0 else diff
        addable = round(addable, 2)
        if addable >= step:
            add = min(addable, diff)
            add = round_step(add, step)
            res[i] += add
            diff -= add
            diff = round(diff, 2)
        i += 1

    # If diff < 0, reduce divisions similarly
    i = 0
    while diff <= -step and i < divisions:
        reducible = res[i]
        if reducible >= step:
            sub = min(reducible, abs(diff))
            sub = round_step(sub, step)
            res[i] -= sub
            diff += sub
            diff = round(diff, 2)
        i += 1

    return res

def distribute_random(total, divisions, max_per_comp, step=0.25):
    # If no cap or cap > total, just generate random fractions scaled
    if max_per_comp <= 0 or max_per_comp >= total:
        vals = [random.random() for _ in range(divisions)]
        s = sum(vals)
        res = [round_step(total * v / s, step) for v in vals]
    else:
        # Generate capped random marks per division
        # Generate random numbers capped at max_per_comp, then scale to total without exceeding caps
        attempts = 0
        max_attempts = 10000
        while attempts < max_attempts:
            vals = [random.uniform(0, max_per_comp) for _ in range(divisions)]
            s = sum(vals)
            if s == 0:
                attempts += 1
                continue
            scale = total / s
            res = [v * scale for v in vals]
            # Check if any exceed cap after scaling
            if all(r <= max_per_comp + 0.01 for r in res):
                res = [round_step(r, step) for r in res]
                break
            attempts += 1
        else:
            # Fallback equal if random can't find solution
            res = distribute_equal(total, divisions, max_per_comp, step)

    # Adjust sum to total by adding or subtracting step carefully
    current_sum = sum(res)
    diff = round(total - current_sum, 2)

    # Try to fix positive diff by adding step to divisions that can add without exceeding cap
    i = 0
    while diff >= step and i < divisions:
        can_add = max_per_comp - res[i] if max_per_comp > 0 else diff
        can_add = round(can_add, 2)
        if can_add >= step:
            add = min(can_add, diff)
            add = round_step(add, step)
            res[i] += add
            diff -= add
            diff = round(diff, 2)
        i += 1

    # Fix negative diff by subtracting step from divisions
    i = 0
    while diff <= -step and i < divisions:
        can_sub = res[i]
        if can_sub >= step:
            sub = min(can_sub, abs(diff))
            sub = round_step(sub, step)
            res[i] -= sub
            diff += sub
            diff = round(diff, 2)
        i += 1

    # Final safety cap check
    res = [min(max_per_comp if max_per_comp > 0 else total, round_step(v, step)) for v in res]
    return res

def process_row(row):
    val = row.get('marks')
    invalid = False
    try:
        val = float(val)
        if val < 0:
            invalid = True
            val = 0
    except:
        invalid = True
        val = 0

    # Check if original marks have decimals
    has_decimal = (val != int(val))
    step = 0.25 if has_decimal else 1.0

    max_comp = max_per_component if max_per_component > 0 else 0

    if division_type == "Equal":
        divisions = distribute_equal(val, num_divisions, max_comp, step)
    else:
        divisions = distribute_random(val, num_divisions, max_comp, step)

    # Convert to int if no decimals in original marks
    if not has_decimal:
        divisions = [int(round(x)) for x in divisions]

    out = {f"div_{i+1}": divisions[i] for i in range(num_divisions)}
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
