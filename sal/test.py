import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO

def generate_docx(data, image_file=None):
    doc = Document()

    # Set font style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Add table for image (right-aligned)
    img_table = doc.add_table(rows=1, cols=2)
    img_cells = img_table.rows[0].cells
    img_cells[0].text = ""
    if image_file:
        run = img_cells[1].paragraphs[0].add_run()
        run.add_picture(image_file, width=Inches(1.0))
    img_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # Ref No. and Date
    ref_table = doc.add_table(rows=1, cols=2)
    ref_table.autofit = False
    ref_table.columns[0].width = Inches(5.5)
    ref_table.columns[1].width = Inches(1.5)

    ref_cells = ref_table.rows[0].cells
    ref_cells[0].text = f"Ref. No.: CUIET/MED/SAL/{data['year']}/{data['semester']}/{data['notice11']}"
    ref_cells[1].text = f"Date: {data['date']}"
    ref_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # Header
    doc.add_paragraph('\nDepartment of ' + data['department']).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph('CUIET – Applied Engineering').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Notice
    doc.add_paragraph('\nNotice').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph('\nSubject: Identification of Slow and Advanced Learners')

    body = (
        f"The below mentioned students of {data['branch']} {data['session']} session July – Dec 2023 were classified "
        "into the Slow and Advanced learners' categories based upon the observations and feedback from the mentors, "
        f"teachers and academic performance in {data['st']} ST-I. Students, who score marks below {data['slow_threshold']}% "
        "are categorized as slow learners and above "
        f"{data['advanced_threshold']}% are categorized as advanced learners. These distinguished parameters enabled in "
        "identification of advanced learners and slow learners. The details of slow and advanced learners is available in Annexure A1."
    )
    doc.add_paragraph('\n' + body)
    doc.add_paragraph('\nNote: Mentors are requested to inform the above students.')

    # Signature block
    sign_table = doc.add_table(rows=1, cols=2)
    sign_cells = sign_table.rows[0].cells
    sign_cells[0].paragraphs[0].add_run("\n\nDean")
    sign_cells[1].paragraphs[0].add_run("\n\nMentor")
    sign_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    doc.add_paragraph('\ncc:\n\n-Notice Board\n\n-Departmental File\n\n-Mentoring File')

    # Save to buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit App
st.title("Notice Generator")

with st.sidebar:
    st.header("Input Fields")
    year = st.text_input("Year", "2025")
    semester = st.text_input("Semester", "02")
    notice11 = st.text_input("Notice No", "011")
    date = st.text_input("Date", "10-04-2025")
    department = st.text_input("Department", "Mechanical Engineering")
    branch = st.text_input("Branch", "ME")
    session = st.text_input("Session", "2022–26")
    st_input = st.text_input("ST", "1st")
    slow_threshold = st.text_input("Slow Learner Threshold (%)", "40")
    advanced_threshold = st.text_input("Advanced Learner Threshold (%)", "75")
    image_file = st.file_uploader("Upload Logo Image", type=["png", "jpg"])

data = {
    "year": year,
    "semester": semester,
    "notice11": notice11,
    "date": date,
    "department": department,
    "branch": branch,
    "session": session,
    "st": st_input,
    "slow_threshold": slow_threshold,
    "advanced_threshold": advanced_threshold
}

if st.button("Generate Notice"):
    docx_file = generate_docx(data, image_file)
    st.success("Notice generated!")
    st.download_button("Download .docx", data=docx_file, file_name="Generated_Notice.docx")