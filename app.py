
import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO
from docx import Document
import base64

# Title and logo
st.markdown("<h1 style='text-align: center; color: red;'>MDH</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center;'>PDF Converter App</h3>", unsafe_allow_html=True)

# Sidebar help section
with st.sidebar.expander("📘 Help & Instructions"):
    st.markdown("""
    **How to Use This App:**
    1. Upload a PDF file.
    2. Choose conversion type: PDF to Word or PDF to Excel.
    3. Click Convert and download the result.

    **Supported Formats:**
    - PDF to Word (.docx)
    - PDF to Excel (.xlsx) with table detection

    **Troubleshooting:**
    - If tables are not detected, OCR will extract text from scanned pages.

    **Contact Support:**
    - Email: support@mdhconverter.com
    """)

# Sidebar version history
with st.sidebar.expander("📄 Version History"):
    st.markdown("""
    #### **Version 1.1.0** — *2024-04-03*
    - Version history section added
    - Improved layout and sidebar navigation
    - Minor bug fixes and performance improvements

    #### **Version 1.0.0** — *2024-04-01*
    - Initial release with PDF to Word and Excel conversion
    - Table detection for Excel
    - MDH branding and logo
    - Help section added
    """)

# File uploader
uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])
conversion_type = st.selectbox("Choose conversion type", ["PDF to Word", "PDF to Excel"])

def convert_pdf_to_word(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    word_doc = Document()
    for page in doc:
        text = page.get_text()
        word_doc.add_paragraph(text)
    output = BytesIO()
    word_doc.save(output)
    return output.getvalue()

def convert_pdf_to_excel(pdf_bytes):
    excel_output = BytesIO()
    writer = pd.ExcelWriter(excel_output, engine='openpyxl')
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        table_found = False
        for i, page in enumerate(pdf.pages):
            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5
            }
            tables = page.extract_tables(table_settings=table_settings)
            if tables:
                table_found = True
                for j, table in enumerate(tables):
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df.to_excel(writer, sheet_name=f"Page{i+1}_Table{j+1}", index=False)
        if not table_found:
            # OCR fallback
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            ocr_text = []
            for page in doc:
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = pytesseract.image_to_string(img)
                ocr_text.append(text)
            df = pd.DataFrame({"OCR_Text": ocr_text})
            df.to_excel(writer, sheet_name="OCR_Text", index=False)
    writer.save()
    return excel_output.getvalue()

if uploaded_file:
    st.success("File uploaded successfully!")
    if st.button("Convert"):
        pdf_bytes = uploaded_file.read()
        if conversion_type == "PDF to Word":
            word_data = convert_pdf_to_word(pdf_bytes)
            st.download_button("Download Word File", word_data, file_name="converted.docx")
        else:
            excel_data = convert_pdf_to_excel(pdf_bytes)
            st.download_button("Download Excel File", excel_data, file_name="converted.xlsx")
