import streamlit as st
import pandas as pd
import numpy as np
import io
import random
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Sidebar setup
st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

# Download sample input
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

# File upload and settings
uploaded_file = st.sidebar.file_uploader("Upload Excel file (1 column: marks)", type=["xlsx"])
num_divisions = st.sidebar.number_input("Number of divisions", min_value=1, max_value=100, value=5)
division_type = st.sidebar.selectbox("Division type", ["Equal", "Random"])
max_per_component_str = st.sidebar.text_input("Max marks per component (optional, comma-separated)", value="")

process_button = st.sidebar.button("Start Processing")

# Title area
st.title("📊 Marks Division Panel")
st.header(":green[Divide 'marks' Column into Divisions]")
st.subheader(":blue[Supports Equal and Random Divisions with Optional Per-Division Max]")

# Parse max caps
def parse_max_caps(text, divisions):
    if not text.strip():
        return [float('inf')] * divisions
    try:
        caps = [float(x.strip()) for x in text.split(',')]
        if len(caps) != divisions:
            return None
        return caps
    except:
        return None

# Split marks into components
def split_marks(total, divisions, mode, caps):
    total = round(float(total), 2)
    result = [0.0] * divisions

    if total == 0:
        return result

    if mode == "Equal":
        per = total / divisions
        for i in range(divisions):
            result[i] = min(round(per, 2), caps[i])
        diff = round(total - sum(result), 2)

        i = 0
        while diff > 0:
            space = caps[i] - result[i]
            if space > 0:
                add = min(space, diff)
                result[i] = round(result[i] + add, 2)
                diff = round(diff - add, 2)
            i = (i + 1) % divisions
    else:
        remaining = total
        while round(remaining, 2) > 0:
            i = random.randint(0, divisions - 1)
            space = caps[i] - result[i]
            if space <= 0:
                continue
            add = round(random.uniform(0.1, min(space, remaining)), 2)
            result[i] += add
            result[i] = round(result[i], 2)
            remaining = round(remaining - add, 2)

    # Final correction
    result = [min(round(x, 2), caps[i]) for i, x in enumerate(result)]
    diff = round(total - sum(result), 2)
    for i in range(divisions):
        if diff <= 0:
            break
        space = caps[i] - result[i]
        if space > 0:
            add = min(space, diff)
            result[i] += add
            result[i] = round(result[i], 2)
            diff -= add

    if float(total).is_integer():
        result = list(map(int, result))
    return result

# Style the output Excel file
def style_output_excel(buffer, highlight_rows):
    wb = load_workbook(buffer)
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

    styled_buffer = io.BytesIO()
    wb.save(styled_buffer)
    styled_buffer.seek(0)
    return styled_buffer

# Process row by row
def process_marks(df, divisions, mode, caps):
    result_data = []
    highlight_rows = []

    for idx, row in df.iterrows():
        mark = row.get('marks')
        try:
            mark = float(mark)
            if mark < 0:
                raise ValueError
        except:
            mark = 0
            highlight_rows.append(idx)

        parts = split_marks(mark, divisions, mode, caps)
        record = {'marks': mark}
        for i, val in enumerate(parts, 1):
            record[f'div_{i}'] = val
        result_data.append(record)

    out_df = pd.DataFrame(result_data)
    buffer = io.BytesIO()
    out_df.to_excel(buffer, index=False)
    buffer.seek(0)
    return style_output_excel(buffer, highlight_rows)

# Run the processing
if process_button and uploaded_file:
    caps = parse_max_caps(max_per_component_str, num_divisions)
    if caps is None:
        st.error("❌ Max marks list must have the same number of values as divisions.")
    else:
        df = pd.read_excel(uploaded_file)
        if 'marks' not in df.columns:
            st.error("❌ Uploaded file must contain a 'marks' column.")
        else:
            result_file = process_marks(df, num_divisions, division_type, caps)
            st.success("✅ File processed successfully.")
            st.download_button(
                label="📥 Download Processed File",
                data=result_file,
                file_name=f"processed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
