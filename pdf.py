import streamlit as st
from msoffice2pdf import convert
from pathlib import Path

def convert_pptx_to_pdf(input_path, output_path):
    convert(input_path, output_path)

def main():
    st.title("PowerPoint to PDF Converter")

    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["ppt", "pptx"])
    if uploaded_file is not None:
        input_path = Path("uploaded_presentation.pptx")
        output_path = Path("converted_presentation.pdf")

        # Save the uploaded file to disk
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Convert the PowerPoint to PDF
        convert_pptx_to_pdf(input_path, output_path)

        # Provide a download link for the converted PDF
        with open(output_path, "rb") as f:
            st.download_button(
                label="Download PDF",
                data=f,
                file_name=output_path.name,
                mime="application/pdf"
            )

if __name__ == "__main__":
    main()
