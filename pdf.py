import streamlit as st
from PIL import Image
import fitz  # PyMuPDF
import io

# Function to extract images from a PDF file given its bytes.
def extract_images_from_pdf(pdf_bytes):
    images = []
    # Open the PDF from bytes (PyMuPDF accepts a stream)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        # Render page to a pixmap
        pix = page.get_pixmap()
        # Convert pixmap to a PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()
    return images

# Function to merge 4 images into a single 2x2 grid.
def merge_images_grid(images, grid_size=(2, 2)):
    # Determine a common size by taking the minimum width and height
    widths, heights = zip(*(img.size for img in images))
    new_width, new_height = min(widths), min(heights)
    resized = [img.resize((new_width, new_height)) for img in images]
    # Create a new blank image with size equal to a 2x2 grid
    grid_width = grid_size[0] * new_width
    grid_height = grid_size[1] * new_height
    merged_img = Image.new("RGB", (grid_width, grid_height), color=(255, 255, 255))
    for idx, img in enumerate(resized):
        row = idx // grid_size[0]
        col = idx % grid_size[0]
        merged_img.paste(img, (col * new_width, row * new_height))
    return merged_img

# Function to convert a PPT/PPTX file to PDF using Aspose.Slides.
def convert_ppt_to_pdf(ppt_bytes):
    import aspose.slides as slides
    # Use a BytesIO stream for the PPT file.
    ppt_io = io.BytesIO(ppt_bytes)
    pres = slides.Presentation(ppt_io)
    out_pdf = io.BytesIO()
    # Save the presentation as PDF.
    pres.save(out_pdf, slides.export.SaveFormat.PDF)
    pres.dispose()
    out_pdf.seek(0)
    return out_pdf.read()

st.title("Merge 4 Pages into One")
st.write(
    "Upload PDF or PPT/PPTX files. The app extracts each page as an image and "
    "merges every 4 pages into a single image (displayed as a 2×2 grid)."
)

# Allow multiple file uploads (PDF, PPT, PPTX)
uploaded_files = st.file_uploader(
    "Choose PDF/PPT files", type=["pdf", "ppt", "pptx"], accept_multiple_files=True
)

all_page_images = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        file_bytes = uploaded_file.read()
        if file_ext in ["ppt", "pptx"]:
            st.info(f"Converting {uploaded_file.name} from PPT to PDF...")
            try:
                pdf_bytes = convert_ppt_to_pdf(file_bytes)
                images = extract_images_from_pdf(pdf_bytes)
                st.success(f"Converted {uploaded_file.name} and extracted {len(images)} pages.")
            except Exception as e:
                st.error(f"Error converting {uploaded_file.name}: {e}")
                images = []
        elif file_ext == "pdf":
            try:
                images = extract_images_from_pdf(file_bytes)
                st.success(f"Extracted {len(images)} pages from {uploaded_file.name}.")
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")
                images = []
        else:
            images = []
        all_page_images.extend(images)
    
    if all_page_images:
        st.write("Merging pages (4 pages per merged image)...")
        merged_images = []
        # Process images in groups of 4.
        for i in range(0, len(all_page_images), 4):
            group = all_page_images[i : i + 4]
            # If fewer than 4 images, append a blank image to fill the grid.
            if len(group) < 4:
                blank = Image.new("RGB", group[0].size, (255, 255, 255))
                while len(group) < 4:
                    group.append(blank)
            merged = merge_images_grid(group)
            merged_images.append(merged)
        
        st.write(f"Created {len(merged_images)} merged images.")
        for idx, img in enumerate(merged_images):
            st.image(img, caption=f"Merged Image {idx + 1}")
            # Prepare image for download.
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            byte_im = buf.getvalue()
            st.download_button(
                label=f"Download Merged Image {idx + 1}",
                data=byte_im,
                file_name=f"merged_{idx + 1}.png",
                mime="image/png",
            )
