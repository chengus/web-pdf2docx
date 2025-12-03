# URL
Use here: [https://web-pdf2docx.streamlit.app/](https://web-pdf2docx.streamlit.app/)

# web-pdf2docx
Web-based pdf to docx converter and translator built with Streamlit and pdf2docx. Why? Because my parents need it.

**Usage (UI)**
- Upload a PDF using the file uploader.
- Optionally set a start page and/or end page (1-based indices). Check the corresponding "Use start page" / "Use end page" boxes to enable them.
- Toggle "Enable multi-processing" to try to speed up conversion for large documents (system-dependent).
- Edit the output filename if you want a custom name, then click `Convert to DOCX`.
- When conversion finishes, click the provided `Download DOCX` button.

**Acknowledgements**
- `streamlit` for the quick UI scaffolding
- `pdf2docx` for the PDF→DOCX conversion engine
