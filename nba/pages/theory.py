import streamlit as st
import pandas as pd
import zipfile
import io
import numpy as np
import base64

# Page headings
st.title("Theory Marks Processing Panel")
st.header(":green[Theory Marks]", divider="rainbow")
st.subheader(":red[ST1, ST2, and ETE marks are equally divided among components based on the number of COs. Two files are generated: one with processed data where all values are valid, and another with unprocessed rows where any value is missing or invalid.]", divider="rainbow")

# Sidebar headings
st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

# Sample input button
if st.sidebar.button("Download Sample Input File"):
    sample_columns = ['sno', 'roll', 'name', 'course-code', 'st1', 'st2', 'ete', 'co']
    sample_df = pd.DataFrame(columns=sample_columns)
    output = io.BytesIO()
    sample_df.to_excel(output, index=False)
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<meta http-equiv="refresh" content="0;url=data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}">'
    st.markdown(href, unsafe_allow_html=True)

uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx"])

def trigger_download(zip_data, filename="output_files.zip"):
    """Trigger automatic download of a zip file via a base64 link."""
    b64 = base64.b64encode(zip_data).decode()
    href = f'<meta http-equiv="refresh" content="0;url=data:application/zip;base64,{b64}">'
    st.markdown(href, unsafe_allow_html=True)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Preview of Uploaded File", df.head())

    required_columns = ['sno', 'roll', 'name', 'course-code', 'st1', 'st2', 'ete', 'co']
    if not all(col in df.columns for col in required_columns):
        st.error("Missing one or more required columns.")
    else:
        processed_rows = []
        unprocessed_rows = []

        for _, row in df.iterrows():
            try:
                if all(pd.notna(row[col]) and str(row[col]).strip() != '' for col in ['st1', 'st2', 'ete', 'co']):
                    st1 = float(row['st1'])
                    st2 = float(row['st2'])
                    ete = float(row['ete'])
                    co = int(row['co'])

                    if co <= 0:
                        raise ValueError("co must be > 0")

                    # Divide and round while preserving total
                    def divide_evenly(value, co):
                        parts = np.round([value / co] * co, 2)
                        parts[-1] = round(value - sum(parts[:-1]), 2)
                        return parts

                    st1_parts = divide_evenly(st1, co)
                    st2_parts = divide_evenly(st2, co)
                    ete_parts = divide_evenly(ete, co)

                    st1_dict = {f'st1_{i+1}': st1_parts[i] for i in range(co)}
                    st2_dict = {f'st2_{i+1}': st2_parts[i] for i in range(co)}
                    ete_dict = {f'ete_{i+1}': ete_parts[i] for i in range(co)}

                    new_row = row.to_dict()
                    new_row.update(st1_dict)
                    new_row.update(st2_dict)
                    new_row.update(ete_dict)

                    processed_rows.append(new_row)
                else:
                    unprocessed_rows.append(row)
            except Exception:
                unprocessed_rows.append(row)

        # DataFrames
        processed_df = pd.DataFrame(processed_rows)
        unprocessed_df = pd.DataFrame(unprocessed_rows)

        # Reorder processed_df
        original_cols = df.columns.tolist()
        st1_cols = sorted([col for col in processed_df.columns if col.startswith('st1_')], key=lambda x: int(x.split('_')[1]))
        st2_cols = sorted([col for col in processed_df.columns if col.startswith('st2_')], key=lambda x: int(x.split('_')[1]))
        ete_cols = sorted([col for col in processed_df.columns if col.startswith('ete_')], key=lambda x: int(x.split('_')[1]))
        ordered_cols = original_cols + st1_cols + st2_cols + ete_cols
        processed_df = processed_df[ordered_cols]

        # Display
        st.write("### ✅ Processed Rows", processed_df.head())
        st.write("### ⚠️ Unprocessed Rows", unprocessed_df.head())

        # Create zip
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as zipf:
            processed_io = io.BytesIO()
            processed_df.to_excel(processed_io, index=False)
            zipf.writestr("processed.xlsx", processed_io.getvalue())

            unprocessed_io = io.BytesIO()
            unprocessed_df.to_excel(unprocessed_io, index=False)
            zipf.writestr("unprocessed.xlsx", unprocessed_io.getvalue())

        buffer.seek(0)
        trigger_download(buffer.getvalue())
