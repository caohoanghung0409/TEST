import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import time
import base64
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

OCR_CONFIG = "--oem 3 --psm 6"

# =========================
# SESSION STATE
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False
if "done" not in st.session_state:
    st.session_state.done = False
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

# =========================
# STYLE (GIỮ NGUYÊN)
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

.file-row {
    margin-top:12px;
    padding:10px;
    border-radius:12px;
    background:white;
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}

.progress-bar {
    height:8px;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="header">🚀 THL PDF → EXCEL</div>', unsafe_allow_html=True)

# =========================
# OCR (FAST + STABLE)
# =========================
def ocr_extract(img):
    w, h = img.size
    crop = img.crop((0, 0, w, int(h * 0.45)))  # chỉ scan vùng cần

    text = pytesseract.image_to_string(
        crop,
        lang="eng",
        config=OCR_CONFIG
    )

    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    if sm and date:
        return sm.group(1), date.group(1)

    return None, None

# =========================
# PROCESS PDF → IMAGES CACHE
# =========================
def load_pdf(pdf_bytes):
    return convert_from_bytes(pdf_bytes, dpi=110)

# =========================
# EXTRACT PER PDF
# =========================
def extract_pdf(images, name, box, global_box, start_time, processed, total):

    results = []
    total_pages = len(images)

    for i, img in enumerate(images, 1):

        processed[0] += 1

        percent = int(i / total_pages * 100)
        global_percent = int(processed[0] / total * 100)

        elapsed = time.time() - start_time
        speed = processed[0] / elapsed if elapsed else 0
        eta = int((total - processed[0]) / speed) if speed else 0

        global_box.markdown(
            f"⚡ {global_percent}% | ⏳ {eta}s",
            unsafe_allow_html=True
        )

        box.markdown(f"""
        <div class="file-row">
            📄 {name} - Trang {i}/{total_pages}
            <div class="progress-bar" style="width:{percent}%"></div>
        </div>
        """, unsafe_allow_html=True)

        sm, date = ocr_extract(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date,
                "Trang": i
            })

    return results

# =========================
# UI UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Chọn PDF",
    type=["pdf"],
    accept_multiple_files=True
)

st.markdown('<div class="header"></div>', unsafe_allow_html=True)

# =========================
# MAIN LOGIC
# =========================
if uploaded_files:

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if st.button("🚀 Bắt đầu xử lý"):

        st.session_state.processing = True

    if st.session_state.processing:

        start_time = time.time()
        processed = [0]

        pdf_list = []
        images_list = []

        # =========================
        # FIX: đọc file 1 lần
        # =========================
        for f in uploaded_files:
            pdf_bytes = f.getvalue()
            pdf_list.append((f.name, pdf_bytes))
            images_list.append(load_pdf(pdf_bytes))

        total_pages = sum(len(x) for x in images_list)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:

            for idx, ((name, _), images) in enumerate(zip(pdf_list, images_list)):

                data = extract_pdf(
                    images,
                    name,
                    boxes[idx],
                    global_box,
                    start_time,
                    processed,
                    total_pages
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))
                    df.to_excel(writer, sheet_name=name[:31], index=False)

        wb = load_workbook(tmp.name)

        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 18

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border

            for cell in ws[1]:
                cell.font = Font(bold=True)

        wb.save(tmp.name)

        st.session_state.excel_file = tmp.name
        st.session_state.processing = False
        st.session_state.done = True

        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH!")

    with open(st.session_state.excel_file, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    st.markdown(f"""
    <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.processing = False
        st.rerun()
