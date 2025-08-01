import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import numpy as np
import cv2

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.set_page_config(page_title="PDF OCR App", layout="wide")

st.title("📄 High-Quality PDF OCR App")
st.markdown("Upload a PDF file, and the app will extract text from all pages using high-resolution OCR.")

uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"])

def pdf_to_images(pdf_bytes, dpi=300):
    images = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
    return images

def preprocess_image(image):
    # Convert to grayscale, apply thresholding (optional but improves OCR)
    img_np = np.array(image)
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    return Image.fromarray(thresh)

if uploaded_file:
    with st.spinner("🔍 Processing PDF and performing OCR..."):
        pdf_bytes = uploaded_file.read()
        images = pdf_to_images(pdf_bytes)

        extracted_text = ""
        for i, image in enumerate(images):
            st.subheader(f"Page {i + 1}")
            preprocessed = preprocess_image(image)
            text = pytesseract.image_to_string(preprocessed, lang='eng')
            extracted_text += f"\n\n--- Page {i+1} ---\n{text}"
            st.image(preprocessed, caption=f"Page {i + 1} (Preprocessed)", use_column_width=True)
            st.text_area(f"OCR Output - Page {i + 1}", text, height=200)

        st.subheader("📄 Combined OCR Text Output")
        st.download_button("📥 Download OCR Text as .txt", extracted_text, file_name="ocr_output.txt")
