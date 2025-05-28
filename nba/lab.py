import streamlit as st
import pandas as pd
import zipfile
import io
import numpy as np
import base64

# Page headings
st.title("Lab Marks Processing Panel")
st.header(":green[Lab Marks]", divider="rainbow")
st.subheader(":red[Internal Viva and External Viva Marks are divided equally among components whose number is based on the number of COs in particular Lab. Two files are generated, one with processed data for which all components are available and other with unprocessed data where at least one of the columns is not available]", divider="rainbow")

# Sidebar headings
st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

# Sample input button
if st.sidebar.button("Download Sample Input File"):
    sample_columns = ['sno', 'roll', 'name', 'course-code', 'l1', 'l2', 'l3', 'l4', 'iv', 'ev', 'co']
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

    required_columns = ['sno', 'roll', 'name', 'course-code', 'l1', 'l2', 'l3', 'l4', 'iv', 'ev', 'co']
    if not all(col in df.columns for col in required_columns):
        st.error("Missing one or more required columns.")
    else:
        processed_rows = []
        unprocessed_rows = []

        for _, row in df.iterrows():
            try:
                if all(pd.notna(row[col]) and str(row[col]).strip() != '' for col in ['l1', 'l2', 'l3', 'l4', 'iv', 'ev', 'co']):
                    iv = float(row['iv'])
                    ev = float(row['ev'])
                    co = int(row['co'])

                    if co <= 0:
                        raise ValueError("co must be > 0")

                    iv_parts = np.round([iv / co] * co, 2)
                    iv_parts[-1] = round(iv - sum(iv_parts[:-1]), 2)

                    ev_parts = np.round([ev / co] * co, 2)
                    ev_parts[-1] = round(ev - sum(ev_parts[:-1]), 2)

                    iv_dict = {f'iv_{i+1}': iv_parts[i] for i in range(co)}
                    ev_dict = {f'ev_{i+1}': ev_parts[i] for i in range(co)}

                    new_row = row.to_dict()
                    new_row.update(iv_dict)
                    new_row.update(ev_dict)

                    processed_rows.append(new_row)
                else:
                    unprocessed_rows.append(row)
            except Exception:
                unprocessed_rows.append(row)

        processed_df = pd.DataFrame(processed_rows)
        unprocessed_df = pd.DataFrame(unprocessed_rows)

        original_cols = df.columns.tolist()
        iv_cols = sorted([col for col in processed_df.columns if col.startswith('iv_')], key=lambda x: int(x.split('_')[1]))
        ev_cols = sorted([col for col in processed_df.columns if col.startswith('ev_')], key=lambda x: int(x.split('_')[1]))
        ordered_cols = original_cols + iv_cols + ev_cols
        processed_df = processed_df[ordered_cols]

        st.write("### ✅ Processed Rows", processed_df.head())
        st.write("### ⚠️ Unprocessed Rows", unprocessed_df.head())

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
