import streamlit as st
import pandas as pd
import zipfile
import io
import numpy as np
import base64

# Page headings
st.title("Assessment Marks Processing Panel")
st.header(":green[IM & EM Marks]", divider="rainbow")
st.subheader(":red[Internal (IM) and External (EM) Marks are equally divided based on the number of COs. Two output files are generated: one with complete processed rows and one with unprocessed rows due to missing/invalid data.]", divider="rainbow")

# Sidebar headings
st.sidebar.title(":rainbow[Dr. Sonu Sharma Apps]")
st.sidebar.subheader("Input/Output")

# Button to download sample input file
if st.sidebar.button("Download Sample Input File"):
    sample_columns = ['sno', 'roll', 'name', 'course-code', 'im', 'em', 'co']
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

    required_columns = ['sno', 'roll', 'name', 'course-code', 'im', 'em', 'co']
    if not all(col in df.columns for col in required_columns):
        st.error("Missing one or more required columns.")
    else:
        processed_rows = []
        unprocessed_rows = []

        for _, row in df.iterrows():
            try:
                if all(pd.notna(row[col]) and str(row[col]).strip() != '' for col in ['im', 'em', 'co']):
                    im = float(row['im'])
                    em = float(row['em'])
                    co = int(row['co'])

                    if co <= 0:
                        raise ValueError("co must be > 0")

                    def divide_evenly(value, co):
                        parts = np.round([value / co] * co, 2)
                        parts[-1] = round(value - sum(parts[:-1]), 2)
                        return parts

                    im_parts = divide_evenly(im, co)
                    em_parts = divide_evenly(em, co)

                    im_dict = {f'im_{i+1}': im_parts[i] for i in range(co)}
                    em_dict = {f'em_{i+1}': em_parts[i] for i in range(co)}

                    new_row = row.to_dict()
                    new_row.update(im_dict)
                    new_row.update(em_dict)

                    processed_rows.append(new_row)
                else:
                    unprocessed_rows.append(row)
            except Exception:
                unprocessed_rows.append(row)

        processed_df = pd.DataFrame(processed_rows)
        unprocessed_df = pd.DataFrame(unprocessed_rows)

        # Reorder columns
        original_cols = df.columns.tolist()
        im_cols = sorted([col for col in processed_df.columns if col.startswith('im_')], key=lambda x: int(x.split('_')[1]))
        em_cols = sorted([col for col in processed_df.columns if col.startswith('em_')], key=lambda x: int(x.split('_')[1]))
        ordered_cols = original_cols + im_cols + em_cols
        processed_df = processed_df[ordered_cols]

        # Display previews
        st.write("### ✅ Processed Rows", processed_df.head())
        st.write("### ⚠️ Unprocessed Rows", unprocessed_df.head())

        # Create ZIP for download
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
