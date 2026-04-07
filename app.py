import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import zipfile
import os
import time
import base64
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# SESSION
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False

if "done" not in st.session_state:
    st.session_state.done = False

if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

if "last_uploaded_names" not in st.session_state:
    st.session_state.last_uploaded_names = []

if "excel" not in st.session_state:
    st.session_state.excel = None

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header {
    font-size:22px;
    font-weight:700;
    margin-bottom:10px;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #93c5fd;
    padding: 25px;
    border-radius: 18px;
    background: white;
}

div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:12px 24px;
    font-weight:600;
    font-size:15px;
    box-shadow:0 4px 14px rgba(0,0,0,0.15);
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 THL PDF → EXCEL </div>', unsafe_allow_html=True)

# =========================
# UPLOAD
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

current_names = [f.name for f in uploaded_files] if uploaded_files else []

if current_names != st.session_state.last_uploaded_names:
    st.session_state.processing = False
    st.session_state.done = False
    st.session_state.last_uploaded_names = current_names

# =========================
# OCR FUNCTION
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# EXTRACT PDF
# =========================
def extract_pdf(file):
    results = []
    images = convert_from_bytes(file.read(), dpi=150)

    for img in images:
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)
        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    return results

# =========================
# MAIN PROCESS
# =========================
if uploaded_files:

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start_time = time.time()

        excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(excel_file.name, engine='openpyxl') as writer:

            for f in uploaded_files:
                data = extract_pdf(f)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        wb = load_workbook(excel_file.name)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

        wb.save(excel_file.name)

        st.session_state.excel = excel_file.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# AUTO DOWNLOAD (HIDDEN + AUTO TRIGGER)
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    with open(st.session_state.excel, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    filename = "TPN_PDF_TO_EXCEL.xlsx"

    download_html = f"""
    <html>
    <body>
        <a id="download" 
           href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
           download="{filename}">
        </a>

        <script>
            document.getElementById('download').click();
        </script>
    </body>
    </html>
    """

    # AUTO DOWNLOAD TRIGGER
    st.components.v1.html(download_html, height=0)

    # HIDE BUTTON (KHÔNG HIỂN THỊ)
    st.markdown("""
        <style>
        .stDownloadButton {display:none !important;}
        </style>
    """, unsafe_allow_html=True)

    # RESET
    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.processing = False
        st.rerun()
