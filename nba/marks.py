import streamlit as st
import pandas as pd
import numpy as np
import io
import random
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("📊 Split ST1, ST2, and ETE Marks | Export as ZIP")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

# Splitter utility
def split_marks(total, structure, na_groups):
    max_cols = list(structure.keys())
    result = {}
    na_choices = [random.choice(group) for group in na_groups]
    allowed_cols = [col for col in max_cols if col not in na_choices]
    max_values = [structure[col] for col in allowed_cols]

    while True:
        assigned = [random.randint(0, mx) for mx in max_values]
        raw_total = sum(assigned)
        if raw_total == 0:
            continue
        scaled = [int(val * total / raw_total) for val in assigned]
        for i, col in enumerate(allowed_cols):
            scaled[i] = min(scaled[i], structure[col])
        diff = total - sum(scaled)
        attempts = 0
        while diff != 0 and attempts < 1000:
            for i, col in enumerate(allowed_cols):
                if diff == 0:
                    break
                if diff > 0 and scaled[i] < structure[col]:
                    scaled[i] += 1
                    diff -= 1
                elif diff < 0 and scaled[i] > 0:
                    scaled[i] -= 1
                    diff += 1
            attempts += 1
        if sum(scaled) == total:
            break

    idx = 0
    for col in max_cols:
        if col in na_choices:
            result[col] = "N/A"
        else:
            result[col] = scaled[idx]
            idx += 1
    return result

# Excel Styling
def style_excel(file_buffer):
    wb = load_workbook(file_buffer)
    ws = wb.active

    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal="center", vertical="center")
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center
            cell.border = thin
            if cell.row == 1:
                cell.fill = header_fill

    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

    styled_io = io.BytesIO()
    wb.save(styled_io)
    styled_io.seek(0)
    return styled_io

# Structures
structure_13 = {
    1: 5,
    2: 2, 3: 2, 4: 2, 5: 2, 6: 2, 7: 2,
    8: 5, 9: 5, 10: 5, 11: 5,
    12: 10, 13: 10
}
structure_16 = {
    'q1': 5,
    'q2': 2, 'q3': 2, 'q4': 2, 'q5': 2, 'q6': 2, 'q7': 2,
    'q8': 5, 'q9': 5, 'q10': 5, 'q11': 5, 'q12': 5, 'q13': 5,
    'q14': 10, 'q15': 10, 'q16': 10
}

# Define NA groups
na_groups_13 = [[2, 3, 4, 5, 6, 7], [8, 9, 10, 11], [12, 13]]
na_groups_16 = [['q2', 'q3', 'q4', 'q5', 'q6', 'q7'],
                ['q8', 'q9', 'q10', 'q11', 'q12', 'q13'],
                ['q14', 'q15', 'q16']]

if uploaded_files:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for uploaded_file in uploaded_files:
            df = pd.read_excel(uploaded_file)

            all_rows = []
            for _, row in df.iterrows():
                st1 = int(row['st1-marks']) if not pd.isna(row['st1-marks']) else 0
                st2 = int(row['st2-marks']) if not pd.isna(row['st2-marks']) else 0
                ete = int(row['ete-marks']) if not pd.isna(row['ete-marks']) else 0

                st1_split = split_marks(st1, structure_13, na_groups_13)
                st2_split = split_marks(st2, structure_13, na_groups_13)
                ete_split = split_marks(ete, structure_16, na_groups_16)

                flat = {
                    **row.to_dict(),
                    **{f"st1-{k}": v for k, v in st1_split.items()},
                    **{f"st2-{k}": v for k, v in st2_split.items()},
                    **{f"ete-{k}": v for k, v in ete_split.items()}
                }
                all_rows.append(flat)

            out_df = pd.DataFrame(all_rows)
            output_io = io.BytesIO()
            out_df.to_excel(output_io, index=False)
            output_io.seek(0)
            styled_output = style_excel(output_io)
            zipf.writestr(f"processed_{uploaded_file.name}", styled_output.read())

    zip_buffer.seek(0)
    st.download_button(
        label="📦 Download All Processed Files as ZIP",
        data=zip_buffer,
        file_name="all_processed_files.zip",
        mime="application/zip"
    )
