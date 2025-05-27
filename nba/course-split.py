import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile

st.set_page_config(page_title="Course Splitter", layout="wide")
st.title("Course Splitter")
st.header(":green[ST1, ST2, ETE File Course Splitter]", divider="rainbow")
st.subheader(":red[Splits the courses in separate excel sheets]", divider="rainbow")
st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

# 1. Auto-download sample input file (no button needed)
sample_data = {
    'id': ['23ME1001', '23ME1002'],
    'name': ['Alice', 'Bob'],
    'course-code': ['24MEC0505', '24MEC0505'],
    **{f'st1-{i}': [i, i+1] for i in range(1, 11)},
    **{f'st2-{i}': [i, i+1] for i in range(1, 11)},
    **{f'ete-q{i}': [i, i+1] for i in range(1, 14)}
}
sample_df = pd.DataFrame(sample_data)

sample_output = BytesIO()
sample_df.to_excel(sample_output, index=False, engine='openpyxl')
sample_output.seek(0)

# Automatically shows the download button on load
st.sidebar.download_button(
    label="📄 Download Sample Input File",
    data=sample_output,
    file_name="Sample_Course_Input.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# 2. File uploader
input_file = st.sidebar.file_uploader("Upload Input Excel File", type=["xlsx"], key="input")

# Define output format columns (Embedded)
output_format_columns = ['Class Roll Number', 'University Roll Number', 'name']
output_format_columns += [f'st1-{i}' for i in range(1, 11)]
output_format_columns += [f'st2-{i}' for i in range(1, 11)]
output_format_columns += [f'ete-q{i}' for i in range(1, 14)]
flat_template = pd.DataFrame(columns=output_format_columns)

# 3. Processing uploaded file
if input_file:
    input_df = pd.read_excel(input_file)
    course_codes = input_df['course-code'].unique()
    output_files = {}

    for course in course_codes:
        course_df = input_df[input_df['course-code'] == course]
        data_rows = []

        for _, row in course_df.iterrows():
            data_row = {
                'Class Roll Number': row['id'],
                'University Roll Number': row['id'],
                'name': row['name']
            }

            for i in range(1, 11):
                data_row[f'st1-{i}'] = row.get(f'st1-{i}', '')
                data_row[f'st2-{i}'] = row.get(f'st2-{i}', '')
            for i in range(1, 14):
                data_row[f'ete-q{i}'] = row.get(f'ete-q{i}', '')

            data_rows.append(data_row)

        output_df = pd.concat([flat_template, pd.DataFrame(data_rows)], ignore_index=True)

        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            output_df.to_excel(writer, index=False)
        output_buffer.seek(0)
        output_files[f"output_{course}.xlsx"] = output_buffer

    # 4. Zip all outputs
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for filename, file_buffer in output_files.items():
            zipf.writestr(filename, file_buffer.getvalue())
    zip_buffer.seek(0)

    # 5. Download ZIP
    st.sidebar.download_button(
        label="📦 Download All Output Files as ZIP",
        data=zip_buffer,
        file_name="all_course_outputs.zip",
        mime="application/zip"
    )
