import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

st.title("📊 ETE Marks Mapper with Highlighting")

highlight_red_ids = ['2010990024', '2055991123', '2055991126', '2055991600']

def clean_col(col):
    return col.fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

def map_marks(super_df, input_df):
    super_id_col = 'id'
    super_course_col = 'course-code'
    input_id_col = 'Admission No. (Roll No.)'
    input_course_col = 'Course Code'

    # Clean merge keys
    super_df[super_id_col] = clean_col(super_df[super_id_col])
    super_df[super_course_col] = clean_col(super_df[super_course_col])
    input_df[input_id_col] = clean_col(input_df[input_id_col])
    input_df[input_course_col] = clean_col(input_df[input_course_col])

    # Sum Q1 parts
    q1_parts = [
        'Obtained Marks Of Q1 \n (a)', 'Obtained Marks Of Q1 \n (b)',
        'Obtained Marks Of Q1 \n (c)', 'Obtained Marks Of Q1 \n (d)',
        'Obtained Marks Of Q1 \n (e)'
    ]
    if all(part in input_df.columns for part in q1_parts):
        input_df['Obtained Marks Of Q1'] = input_df[q1_parts].sum(axis=1)

    # Add Q2 if present
    available_questions = []
    for i in range(1, 17):
        col = f'Obtained Marks Of Q{i}'
        if col in input_df.columns:
            available_questions.append(i)

    # Required columns
    input_df = input_df[[input_id_col, input_course_col] + [f'Obtained Marks Of Q{i}' for i in available_questions]]

    # Merge
    merged_df = pd.merge(
        super_df,
        input_df,
        how='left',
        left_on=[super_id_col, super_course_col],
        right_on=[input_id_col, input_course_col]
    )

    # Map marks to ete-q1 to ete-q16
    for i in available_questions:
        src_col = f'Obtained Marks Of Q{i}'
        dest_col = f'ete-q{i}'
        if dest_col in merged_df.columns:
            merged_df[dest_col] = merged_df[src_col]

    # Drop helper columns
    drop_cols = [input_id_col, input_course_col] + [f'Obtained Marks Of Q{i}' for i in available_questions]
    merged_df.drop(columns=drop_cols, inplace=True, errors='ignore')

    # Ensure all ete-q1 to ete-q16 exist
    for i in range(1, 17):
        col = f'ete-q{i}'
        if col not in merged_df.columns:
            merged_df[col] = pd.NA

    # Replace 0 with 'U'
    for i in range(1, 17):
        col = f'ete-q{i}'
        if col in merged_df.columns:
            merged_df[col] = merged_df[col].replace(0, 'U')

    # Highlight missing marks
    def highlight_rule(row):
        if all(pd.isna(row[f'ete-q{i}']) for i in range(1, 17)):
            return 'red' if row[super_id_col] in highlight_red_ids else 'yellow'
        return ''
    merged_df['__highlight__'] = merged_df.apply(highlight_rule, axis=1)

    return merged_df

def export_with_highlight(df, filename="Mapped_ETE_Marks_Highlighted.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mapped Marks"

    headers = df.drop(columns='__highlight__').columns.tolist()
    for col_idx, col in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=col)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for i, row in enumerate(df.itertuples(index=False), start=2):
        row_data = row[:-1]
        highlight = row[-1]

        for j, val in enumerate(row_data, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            if highlight == 'red':
                cell.fill = red_fill
            elif highlight == 'yellow':
                cell.fill = yellow_fill

    wb.save(filename)
    return filename

# Upload section
super_file = st.file_uploader("📁 Upload Super File (Excel)", type=["xlsx"])
input_files = st.file_uploader("📂 Upload One or More Input Files", type=["xlsx"], accept_multiple_files=True)

if super_file and input_files:
    try:
        # Load super file with correct dtype
        super_df = pd.read_excel(super_file, dtype=str)
        # Load and combine all input files
        input_df_list = []
        for file in input_files:
            df = pd.read_excel(file, dtype=str)
            input_df_list.append(df)
        input_df = pd.concat(input_df_list, ignore_index=True)

        # Check for required columns
        if 'id' not in super_df.columns or 'course-code' not in super_df.columns:
            st.error("❌ Super file must contain 'id' and 'course-code'.")
        elif 'Admission No. (Roll No.)' not in input_df.columns or 'Course Code' not in input_df.columns:
            st.error("❌ Input file(s) must contain 'Admission No. (Roll No.)' and 'Course Code'.")
        else:
            final_df = map_marks(super_df, input_df)

            st.subheader("✅ Mapped & Highlighted Preview")
            st.dataframe(final_df.drop(columns='__highlight__'), use_container_width=True)

            filename = "Mapped_ETE_Marks_Highlighted.xlsx"
            export_with_highlight(final_df, filename)

            with open(filename, "rb") as f:
                st.download_button("⬇️ Download Highlighted Excel", f, file_name=filename)

    except Exception as e:
        st.error(f"❌ Error occurred: {e}")

else:
    st.info("👆 Please upload the Super File and at least one Input File.")
