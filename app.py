# app.py
import io
import tempfile
from pathlib import Path

import streamlit as st
from pdf2docx import Converter


def pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    start_page: int | None = None,
    end_page: int | None = None,
    multi_processing: bool = False,
):
    cv = Converter(pdf_path)

    # pdf2docx uses zero-based page indices
    convert_kwargs = {}
    if start_page is not None:
        convert_kwargs["start"] = max(start_page - 1, 0)
    if end_page is not None:
        # end is inclusive index in pdf2docx
        convert_kwargs["end"] = max(end_page - 1, 0)

    cv.convert(docx_path, multi_processing=multi_processing, **convert_kwargs)
    cv.close()


def main():
    st.title("PDF → DOCX Converter")

    st.markdown(
        "Upload a PDF, adjust options, and download the converted DOCX file."
    )

    uploaded_file = st.file_uploader(
        "Upload PDF file",
        type=["pdf"],
        accept_multiple_files=False,
    )

    # Conversion options
    st.subheader("Conversion options")

    col1, col2 = st.columns(2)
    with col1:
        start_page = st.number_input(
            "Start page (1-based, optional)",
            min_value=1,
            value=1,
            step=1,
        )
        use_start = st.checkbox("Use start page", value=False)
    with col2:
        end_page = st.number_input(
            "End page (1-based, optional)",
            min_value=1,
            value=1,
            step=1,
        )
        use_end = st.checkbox("Use end page", value=False)

    multi_processing = st.checkbox(
        "Enable multi-processing (faster on large PDFs)", value=False
    )

    default_output_name = "output.docx"
    if uploaded_file is not None:
        name_stem = Path(uploaded_file.name).stem
        default_output_name = f"{name_stem}.docx"

    output_filename = st.text_input(
        "Output DOCX filename", value=default_output_name
    )

    # Convert button
    convert_btn = st.button("Convert to DOCX", disabled=(uploaded_file is None))

    if convert_btn and uploaded_file is not None:
        if not output_filename.lower().endswith(".docx"):
            output_filename += ".docx"

        # Ensure page range is consistent
        sp = start_page if use_start else None
        ep = end_page if use_end else None
        if sp is not None and ep is not None and ep < sp:
            st.error("End page cannot be smaller than start page.")
            return

        with st.spinner("Converting..."):
            # Write uploaded PDF to a temporary file
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                tmp_pdf_path = tmp_pdf.name

            # Prepare temp output DOCX path
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
                tmp_docx_path = tmp_docx.name

            # Run conversion
            pdf_to_docx(
                pdf_path=tmp_pdf_path,
                docx_path=tmp_docx_path,
                start_page=sp,
                end_page=ep,
                multi_processing=multi_processing,
            )

            # Read DOCX bytes to serve for download
            with open(tmp_docx_path, "rb") as f:
                docx_bytes = f.read()

        st.success("Conversion complete.")

        st.download_button(
            label="Download DOCX",
            data=docx_bytes,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )


if __name__ == "__main__":
    main()
