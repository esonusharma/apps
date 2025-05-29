import streamlit as st
import pandas as pd
import numpy as np
import io
import random
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Sidebar
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

uploaded_file = st.sidebar.file_uploader("Upload Excel file (1 column: marks)", type=["xlsx"])

num_divisions = st.sidebar.number_input("Number of divisions", min_value=1, max_value=100, value=5)
division_type = st.sidebar.selectbox("Division type", ["Equal", "Random"])
max_per_component = st.sidebar.number_input("Max marks per component (optional)", min_value=0.0, value=10.0, step=0.5)

process_button = st.sidebar.button("Start Processing")

st.title("📊 Marks Division Panel")
st.header(":green[Divide 'marks' Column into Divisions]")
st.subheader(":blue[Supports Equal and Random Divisions with Decimal/Integer Precision]")

# Core logic
def split_single_column_marks(total, divisions, division_type, max_component):
    total = round(float(total), 2)

    if division_type == "Equal":
        base = round(total / divisions, 2)
        result = [min(base, max_component) if max_component > 0 else base for _ in range(divisions)]

        current_sum = round(sum(result), 2)
        diff = round(total - current_sum, 2)

        for i in range(divisions):
            if diff <= 0:
                break
            addable = min(max_component - result[i], diff) if max_component > 0 else diff
            addable = round(addable, 2)
            if addable > 0:
                result[i] = round(result[i] + addable, 2)
                diff = round(diff - addable, 2)

    else:  # Random division
        if total == 0:
            return [0.0] * divisions

        result = [0.0] * divisions
        remaining = total
        max_c = max_component if max_component > 0 else total

        while round(remaining, 2) > 0:
            i = random.randint(0, divisions - 1)
            space = max_c - result[i]
            if space <= 0:
                continue
            add = round(random.uniform(0.1, min(remaining, space)), 2)
            result[i] = round(result[i] + add, 2)
            remaining = round(remaining - add, 2)

    # Final adjustment
    result = [round(min(v, max_component), 2) if max_component > 0 else round(v, 2) for v in result]
    final_sum = round(sum(result), 2)
    diff = round(total - final_sum, 2)

    for i in range(divisions):
        if diff == 0:
            break
        space = max_component - result[i] if max_component > 0 else diff
        if space <= 0:
            continue
        add = round(min(diff, space), 2)
        result[i] = round(result[i] + add, 2)
        diff = round(diff - add, 2)

    # Integer result if input is integer
    result = [round(v, 2) for v in result]
    if float(total).is_integer():
        result = list(map(int, result))

    return result

def style_excel(file_buffer, highlight_rows):
    wb = load_workbook(file_buffer)
    ws = wb.active

    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal="center", vertical="center")
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center
            cell.border = thin
            if cell.row == 1:
                cell.fill = yellow
            elif cell.row - 2 in highlight_rows:
                cell.fill = red

    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

    styled_io = io.BytesIO()
    wb.save(styled_io)
    styled_io.seek(0)
    return styled_io

def process_row(row):
    out = {}
    invalid = False
    val = row.get('marks')
    try:
        val = float(val)
        if val < 0:
            invalid = True
    except:
        invalid = True
        val = 0

    divisions_values = split_single_column_marks(val, num_divisions, division_type, max_per_component)
    for i, v in enumerate(divisions_values, start=1):
        out[f"div_{i}"] = v

    out['marks'] = val
    return out, invalid

def process_file(file):
    df = pd.read_excel(file)
    if 'marks' not in df.columns:
        st.warning(f"File does not contain 'marks' column.")
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
    styled_output = style_excel(buffer, invalid_rows)
    return styled_output

# Process and auto download
if process_button and uploaded_file:
    result_file = process_file(uploaded_file)
    if result_file:
        st.success("✅ File processed successfully. Your download will start below:")
        st.download_button(
            label="📥 Download Processed File",
            data=result_file,
            file_name=f"processed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
