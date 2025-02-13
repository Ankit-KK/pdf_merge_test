import streamlit as st
from msoffice2pdf import convert
from pathlib import Path
import tempfile

def convert_pptx_to_pdf(input_path, output_path, use_libreoffice=False):
    """
    Converts a PowerPoint file to PDF.

    Parameters:
    - input_path: Path to the input PowerPoint file.
    - output_path: Path where the output PDF will be saved.
    - use_libreoffice: Boolean flag to use LibreOffice for conversion.
                       Set to True if Microsoft Office is not available.
    """
    # 'soft' parameter: 0 for Microsoft Office, 1 for LibreOffice
    soft = 1 if use_libreoffice else 0
    convert(source=input_path, output_dir=output_path.parent, soft=soft)

def main():
    st.title("PowerPoint to PDF Converter")

    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["ppt", "pptx"])
    if uploaded_file is not None:
        # Save the uploaded file to a temporary directory
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
            temp_input.write(uploaded_file.getbuffer())
            input_path = Path(temp_input.name)

        output_path = input_path.with_suffix(".pdf")

        # Convert the PowerPoint to PDF
        try:
            convert_pptx_to_pdf(input_path, output_path, use_libreoffice=True)

            # Provide a download link for the converted PDF
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Converted PDF",
                    data=f,
                    file_name=output_path.name,
                    mime="application/pdf"
                )

            st.success("Conversion successful! Click the button above to download the PDF.")

        except Exception as e:
            st.error(f"An error occurred during conversion: {e}")

        finally:
            # Clean up temporary files
            input_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

if __name__ == "__main__":
    main()
