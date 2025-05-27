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

input_file = st.sidebar.file_uploader("Upload Input Excel File", type=["xlsx"], key="input")

# Define output format columns (Embedded)
output_format_columns = ['Class Roll Number', 'University Roll Number', 'name']
output_format_columns += [f'st1-{i}' for i in range(1, 11)]
output_format_columns += [f'st2-{i}' for i in range(1, 11)]
output_format_columns += [f'ete-q{i}' for i in range(1, 14)]
flat_template = pd.DataFrame(columns=output_format_columns)

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

            # Map st1-1 to st1-10
            for i in range(1, 11):
                col = f'st1-{i}'
                if col in row:
                    data_row[col] = row[col]

            # Map st2-1 to st2-10
            for i in range(1, 11):
                col = f'st2-{i}'
                if col in row:
                    data_row[col] = row[col]

            # Map ete-q1 to ete-q13
            for i in range(1, 14):
                col = f'ete-q{i}'
                if col in row:
                    data_row[col] = row[col]

            data_rows.append(data_row)

        # Final output DataFrame
        output_df = pd.concat([flat_template, pd.DataFrame(data_rows)], ignore_index=True)

        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            output_df.to_excel(writer, index=False)
        output_buffer.seek(0)
        output_files[f"output_{course}.xlsx"] = output_buffer

    # Zip all output files
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for filename, file_buffer in output_files.items():
            zipf.writestr(filename, file_buffer.getvalue())
    zip_buffer.seek(0)

    # Single download button for ZIP in sidebar
    st.sidebar.download_button(
        label="📦 Download All Output Files as ZIP",
        data=zip_buffer,
        file_name="all_course_outputs.zip",
        mime="application/zip"
    )
