import streamlit as st
from PIL import Image
import fitz  # PyMuPDF
import io

# Set page config according to Streamlit guidelines.
st.set_page_config(
    page_title="Merge 4 Pages into One",
    page_icon=":page_facing_up:",
    layout="wide"
)

# Cache the function to extract images from a PDF.
@st.cache_data
def extract_images_from_pdf(pdf_bytes):
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        # Render page to a pixmap
        pix = page.get_pixmap()
        # Convert the pixmap to a PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()
    return images

# Function to merge 4 images into a 2x2 grid using Pillow.
def merge_images_grid(images, grid_size=(2, 2)):
    # Use the minimum dimensions to standardize each image's size.
    widths, heights = zip(*(img.size for img in images))
    new_width, new_height = min(widths), min(heights)
    resized = [img.resize((new_width, new_height)) for img in images]
    grid_width = grid_size[0] * new_width
    grid_height = grid_size[1] * new_height
    merged_img = Image.new("RGB", (grid_width, grid_height), color=(255, 255, 255))
    for idx, img in enumerate(resized):
        row = idx // grid_size[0]
        col = idx % grid_size[0]
        merged_img.paste(img, (col * new_width, row * new_height))
    return merged_img

# Cache the PPT to PDF conversion since it can be time‐consuming.
@st.cache_data
def convert_ppt_to_pdf(ppt_bytes):
    import aspose.slides as slides
    ppt_io = io.BytesIO(ppt_bytes)
    out_pdf = io.BytesIO()
    # Use a context manager as recommended to ensure proper cleanup.
    with slides.Presentation(ppt_io) as pres:
        pres.save(out_pdf, slides.export.SaveFormat.PDF)
    out_pdf.seek(0)
    return out_pdf.read()

# Streamlit UI starts here.
st.title("Merge 4 Pages into One")
st.write(
    "Upload PDF or PPT/PPTX files. The app extracts each page as an image and merges every 4 pages "
    "into a single image arranged in a 2×2 grid."
)

# File uploader widget (supports multiple files).
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
        # Group pages in sets of 4
        for i in range(0, len(all_page_images), 4):
            group = all_page_images[i : i + 4]
            # If fewer than 4 images remain, add a blank image to complete the grid.
            if len(group) < 4:
                blank = Image.new("RGB", group[0].size, (255, 255, 255))
                while len(group) < 4:
                    group.append(blank)
            merged = merge_images_grid(group)
            merged_images.append(merged)
        
        st.write(f"Created {len(merged_images)} merged image(s).")
        for idx, img in enumerate(merged_images):
            st.image(img, caption=f"Merged Image {idx + 1}")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            byte_im = buf.getvalue()
            st.download_button(
                label=f"Download Merged Image {idx + 1}",
                data=byte_im,
                file_name=f"merged_{idx + 1}.png",
                mime="image/png"
            )
