import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
import pandas as pd
import re
import tempfile
import time
import base64
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# SESSION
# =========================
if "done" not in st.session_state:
    st.session_state.done = False
if "file_path" not in st.session_state:
    st.session_state.file_path = None

# =========================
# STYLE (GIỮ GIAO DIỆN CŨ)
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
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="header">🚀 THL PDF → EXCEL</div>', unsafe_allow_html=True)

# =========================
# REGEX
# =========================
SM_REGEX = re.compile(r"(SM\d{4}\.\d{4})")
DATE_REGEX = re.compile(r"(\d{2}/\d{2}/\d{4})")

# =========================
# OCR
# =========================
def ocr(img):
    text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")
    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)
    if sm and date:
        return sm.group(1), date.group(1)
    return None, None

# =========================
# PROCESS PDF
# =========================
def process(files, progress_box):

    wb = Workbook()
    wb.remove(wb.active)

    global_count = 0

    total_pages = 0
    file_bytes = []

    for f in files:
        b = f.getvalue()
        file_bytes.append((f.name, b))
        total_pages += int(pdfinfo_from_bytes(b)["Pages"])

    start = time.time()

    for name, b in file_bytes:

        ws = wb.create_sheet(name[:31])

        pages = convert_from_bytes(b, dpi=110)

        ws.append(["STT", "SM", "Ngày", "Trang"])

        for i, img in enumerate(pages, start=1):

            global_count += 1

            percent = int(global_count / total_pages * 100)

            elapsed = time.time() - start
            speed = global_count / elapsed if elapsed else 0
            eta = int((total_pages - global_count) / speed) if speed else 0

            progress_box.markdown(f"""
            <div style="background:#ddd;height:18px;border-radius:10px;">
                <div style="width:{percent}%;height:100%;
                background:linear-gradient(90deg,#3b82f6,#22c55e);"></div>
            </div>
            <p>⚡ {percent}% | ETA {eta}s</p>
            """, unsafe_allow_html=True)

            sm, date = ocr(img)

            if sm and date:
                ws.append([i, sm, date, i])

    # format
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                c.border = border
                if row[0].row == 1:
                    c.font = Font(bold=True)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)

    return tmp.name

# =========================
# UI
# =========================
files = st.file_uploader("📂 Chọn PDF", type=["pdf"], accept_multiple_files=True)

progress_box = st.empty()

if files and st.button("🚀 Bắt đầu xử lý"):

    path = process(files, progress_box)

    st.session_state.file_path = path
    st.session_state.done = True
    st.rerun()

# =========================
# DOWNLOAD AUTO
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    with open(st.session_state.file_path, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    # =========================
    # AUTO DOWNLOAD (JS CLICK)
    # =========================
    download_html = f"""
    <a id="download_link"
       href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
       download="output.xlsx"></a>

    <script>
        document.getElementById('download_link').click();
    </script>
    """

    st.markdown(download_html, unsafe_allow_html=True)

    st.download_button(
        "📥 Download lại file",
        data=data,
        file_name="output.xlsx"
    )
