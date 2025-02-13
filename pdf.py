import streamlit as st
from PyPDF2 import PdfReader, PdfWriter, PageObject
from io import BytesIO
import tempfile
import os
from spire.presentation import Presentation
from spire.presentation.common import *

def ppt_to_images(file_path):
    pres = Presentation()
    pres.LoadFromFile(file_path)
    images = []
    for i, slide in enumerate(pres.Slides):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
            image = slide.SaveAsImage()
            image.Save(tmp_img.name)
            images.append(tmp_img.name)
            image.Dispose()
    pres.Dispose()
    return images

def images_to_pdf(images):
    pdf_buffer = BytesIO()
    writer = PdfWriter()
    
    for img_path in images:
        img = PdfReader(img_path)
        writer.add_page(img.pages[0])
        os.remove(img_path)  # Clean up temporary image
    
    writer.write(pdf_buffer)
    pdf_buffer.seek(0)
    return pdf_buffer

def merge_4_pages(pages):
    writer = PdfWriter()
    for i in range(0, len(pages), 4):
        merged_page = PageObject.create_blank_page(
            width=pages[0].mediabox[2] * 2,
            height=pages[0].mediabox[3] * 2
        )
        
        for j in range(4):
            if i + j >= len(pages):
                break
            page = pages[i + j]
            x = (j % 2) * page.mediabox[2]
            y = (1 - (j // 2)) * page.mediabox[3]
            merged_page.merge_translated_page(page, x, y)
        
        writer.add_page(merged_page)
    return writer

# Streamlit UI
st.title("PPT/PDF Merger (4 Pages/Slides per Sheet)")
uploaded_files = st.file_uploader("Upload files", type=["pdf", "pptx"], accept_multiple_files=True)

if st.button("Process Files"):
    if not uploaded_files:
        st.warning("Please upload files first!")
    else:
        all_pages = []
        
        for file in uploaded_files:
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                tmp_file.write(file.read())
                
                if file.name.endswith(".pdf"):
                    pdf = PdfReader(tmp_file.name)
                    all_pages.extend(pdf.pages)
                elif file.name.endswith(".pptx"):
                    images = ppt_to_images(tmp_file.name)
                    pdf_buffer = images_to_pdf(images)
                    pdf = PdfReader(pdf_buffer)
                    all_pages.extend(pdf.pages)
                
                os.remove(tmp_file.name)
        
        writer = merge_4_pages(all_pages)
        merged_pdf = BytesIO()
        writer.write(merged_pdf)
        merged_pdf.seek(0)
        
        st.download_button(
            "Download Merged PDF",
            merged_pdf,
            "merged_output.pdf",
            "application/pdf"
        )
