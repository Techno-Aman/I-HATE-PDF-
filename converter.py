import subprocess
import streamlit as st
from pdf2docx import Converter
from io import BytesIO
import zipfile
import tempfile
import os

# ------------------------------
# Helper conversion functions
# ------------------------------

st.set_page_config(page_title="I💔PDF", layout="wide")

st.markdown("""
<style>

/* App background */
.stApp{
    background-color:black;
    color:white;
}

/* Card buttons */
div.stButton > button {
    background-color:#1f1f1f;
    color:white;
    border-radius:12px;
    padding:14px 18px;
    font-size:15px;
    border:2px solid transparent;
    transition: all 0.25s ease;
}

/* Hover glow */
div.stButton > button:hover {
    border:2px solid #00ADB5;
    box-shadow:0px 0px 12px #00ADB5;
}

/* Selected card */
.selected {
    background-color:white !important;
    color:black !important;
    border:2px solid #00ADB5 !important;
}

/* Convert button */
.convert button {
    background-color:#2a2a2a !important;
    color:white !important;
    border-radius:10px;
    padding:10px 18px;
}

.convert button:hover{
    border:2px solid #00ADB5;
    box-shadow:0px 0px 10px #00ADB5;
}

</style>
""", unsafe_allow_html=True)

st.markdown(
"""
<h1 style='text-align:center;'>I💔PDF</h1>
<p style='text-align:center;'>Convert PDF, Word, and PowerPoint easily</p>
""",
unsafe_allow_html=True
)

st.markdown("<br><br>", unsafe_allow_html=True)
if "conversion_type" not in st.session_state:
    st.session_state.conversion_type = None

col1, col2 = st.columns(2)

with col1:
    if st.button("📄 PDF → Word", use_container_width=True):
        st.session_state.conversion_type = "PDF TO WORD"

with col2:
     if st.button("📝 Word → PDF", use_container_width=True):
        st.session_state.conversion_type = "WORD TO PDF"
        
if st.session_state.conversion_type:
    col1, col2, col3 = st.columns([1,2,1])

    with col2:
        st.info(f"Selected conversion: {st.session_state.conversion_type}")



def pdf_to_docx(pdf_bytes: bytes) -> BytesIO:
    # Create temp PDF file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_bytes)
        pdf_path = temp_pdf.name

    # Output DOCX path
    docx_path = pdf_path.replace(".pdf", ".docx")

    try:
        # Convert PDF → DOCX
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=20)
        cv.close()

        # Read result
        output = BytesIO()
        with open(docx_path, "rb") as f:
            output.write(f.read())

        output.seek(0)
        return output

    finally:
        # Cleanup (VERY IMPORTANT for deployment)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        if os.path.exists(docx_path):
            os.remove(docx_path)


def docx_to_pdf(docx_bytes: bytes) -> BytesIO:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
        temp_docx.write(docx_bytes)
        docx_path = temp_docx.name

    output_dir = tempfile.gettempdir()

    try:
        result = subprocess.run(
                [
                    r"C:\Program Files\LibreOffice\program\soffice.exe",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", output_dir,
                    docx_path
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )

        pdf_path = docx_path.replace(".docx", ".pdf")

        # Check if conversion failed
        if not os.path.exists(pdf_path):
            raise Exception(result.stderr.decode())

        output = BytesIO()
        with open(pdf_path, "rb") as f:
            output.write(f.read())

        output.seek(0)
        return output

    finally:
        if os.path.exists(docx_path):
            os.remove(docx_path)
        if os.path.exists(docx_path.replace(".docx", ".pdf")):
            os.remove(docx_path.replace(".docx", ".pdf"))

# ------------------------------
# Streamlit UI
# ------------------------------

st.markdown("<br><br>", unsafe_allow_html=True)
conversion = st.session_state.conversion_type
colA, colB, colC = st.columns([1,2,1])

with colB:
    if conversion == "PDF TO WORD":
        allowed_types = ["pdf"]
    elif conversion == "WORD TO PDF":
        allowed_types = ["docx"]
    else:
        allowed_types = ["pdf", "docx"]  # default before selection

    uploaded_files = st.file_uploader(
        "Upload your files",
        accept_multiple_files=True,
        type=allowed_types
    )


st.markdown('<div class="convert">', unsafe_allow_html=True)

colx, coly, colz = st.columns([2,1,2])
with coly:
    process_button = st.button("Convert File", use_container_width=True)

st.markdown('</div>', unsafe_allow_html=True)

# warning if card not selected
if process_button and not conversion:
    st.warning("Please select a conversion type first.")
if process_button and uploaded_files and conversion:

    results = []

    for up in uploaded_files:
        name = up.name
        data = up.read()

        filename = None
        out = None

        if conversion == "PDF TO WORD":
            out = pdf_to_docx(data)
            filename = name.replace(".pdf", ".docx")

        elif conversion == "WORD TO PDF":
            out = docx_to_pdf(data)
            filename = name.replace(".docx", ".pdf")

        if filename and out:
            results.append((filename, out.getvalue()))
    # SINGLE FILE DOWNLOAD
    if len(results) == 1:

        col1, col2, col3 = st.columns([1,2,1])

        with col2:
            download_clicked = st.download_button(
                "Download File",
                data=results[0][1],
                file_name=results[0][0]
            )

        if download_clicked:
            st.session_state.conversion_type = None

    # MULTIPLE FILES ZIP
    else:

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for filename, data in results:
                 zf.writestr(filename, data)

        col1, col2, col3 = st.columns([1,2,1])

        with col2:
            download_clicked = st.download_button(
                "Download ZIP",
                data=zip_buffer.getvalue(),
                file_name="converted_files.zip"
            )

        if download_clicked:
            st.session_state.conversion_type = None