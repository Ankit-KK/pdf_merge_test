import streamlit as st
from PyPDF2 import PdfReader, PdfWriter, PageObject
from io import BytesIO
from pptx import Presentation
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import Image
import tempfile
import os

# Helper function to convert PPT slides to images
def ppt_to_images(file):
    prs = Presentation(file)
    images = []
    for i, slide in enumerate(prs.slides):
        img_path = f"slide_{i}.png"
        slide.save_as_image(img_path)
        images.append(img_path)
    return images

# Helper function to convert images to PDF pages
def images_to_pdf(images):
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    for img_path in images:
        img = Image.open(img_path)
        img_width, img_height = img.size
        c.setPageSize((img_width, img_height))
        c.drawImage(img_path, 0, 0, width=img_width, height=img_height)
        c.showPage()
        img.close()
        os.remove(img_path)  # Clean up temporary image files
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer

# Helper function to merge 4 pages into one
def merge_4_pages(pages):
    writer = PdfWriter()
    for i in range(0, len(pages), 4):
        # Create a blank page (adjust dimensions as needed)
        merged_page = PageObject.create_blank_page(
            width=pages[0].mediabox[2] * 2,  # Double width
            height=pages[0].mediabox[3] * 2   # Double height
        )
        
        # Add pages in a 2x2 grid
        for j in range(4):
            if i + j >= len(pages):
                break
            page = pages[i + j]
            x = (j % 2) * page.mediabox[2]
            y = (1 - (j // 2)) * page.mediabox[3]
            merged_page.merge_translated_page(page, x, y)
        
        writer.add_page(merged_page)
    return writer

# Streamlit app
st.title("Merge 4 Pages into One")
st.write("Upload PDF or PPT files to merge 4 pages/slides into a single page.")

uploaded_files = st.file_uploader(
    "Choose files", 
    type=["pdf", "pptx"], 
    accept_multiple_files=True
)

if st.button("Merge Files"):
    if len(uploaded_files) == 0:
        st.warning("Upload at least one file.")
    else:
        all_pages = []
        
        # Process uploaded files
        for file in uploaded_files:
            if file.type == "application/pdf":
                # Extract pages from PDF
                pdf = PdfReader(file)
                all_pages.extend(pdf.pages)
            elif file.type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
                # Convert PPT to images and then to PDF
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
                    tmp_file.write(file.read())
                    tmp_file_path = tmp_file.name
                
                images = ppt_to_images(tmp_file_path)
                pdf_buffer = images_to_pdf(images)
                pdf = PdfReader(pdf_buffer)
                all_pages.extend(pdf.pages)
                os.remove(tmp_file_path)  # Clean up temporary PPT file
        
        # Merge pages into a 2x2 grid
        writer = merge_4_pages(all_pages)
        
        # Save merged PDF to BytesIO buffer
        merged_pdf = BytesIO()
        writer.write(merged_pdf)
        merged_pdf.seek(0)
        
        # Download button for the merged PDF
        st.download_button(
            label="Download Merged PDF",
            data=merged_pdf,
            file_name="merged_output.pdf",
            mime="application/pdf"
        )
