import streamlit as st
from docx import Document
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar, LTTextLine

def get_fonts_from_docx(docx_file):
    doc = Document(docx_file)
    fonts = set()

    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.name:
                fonts.add(run.font.name)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)

    return fonts

def get_fonts_from_pdf(pdf_file):
    fonts = set()
    for page_layout in extract_pages(pdf_file):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for text_line in element:
                    if isinstance(text_line, LTTextLine):
                        for character in text_line:
                            if isinstance(character, LTChar):
                                if character.fontname:
                                    fonts.add(character.fontname)
    return fonts

st.title("Font Types in Document")

uploaded_file = st.file_uploader("Upload a Word or PDF document", type=["docx", "pdf"])

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()
    fonts = set()
    document_type = ""

    if file_extension == "docx":
        fonts = get_fonts_from_docx(uploaded_file)
        document_type = "Word"
    elif file_extension == "pdf":
        fonts = get_fonts_from_pdf(uploaded_file)
        document_type = "PDF"
    else:
        st.error("Unsupported file type. Please upload a DOCX or PDF file.")

    if fonts:
        st.write(f"Fonts used in the {document_type} document:")
        for font in fonts:
            st.write(font)
    elif document_type: # If document_type is set, but no fonts found
        st.write(f"No fonts found in the {document_type} document.")

