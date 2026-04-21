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
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header { font-size:22px; font-weight:700; margin-bottom:10px; }

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
}

.progress {
    height:8px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
    margin-top:6px;
}
.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}

.global-bar {
    height:20px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
    margin-top:10px;
}
.global-fill {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 THL PDF → EXCEL </div>', unsafe_allow_html=True)

# =========================
# UPLOADER
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
# OCR (FIX MISS + SAI KÝ TỰ)
# =========================
def ocr_extract(img):

    def normalize(text):
        text = text.replace("P8", "PR")
        text = text.replace("S0", "SO")
        text = text.replace(" ", "")
        return text

    def read_lines(image):

        texts = []
        for psm in [6, 11]:
            t = pytesseract.image_to_string(
                image,
                lang='eng',
                config=f'--oem 3 --psm {psm}'
            )
            texts.append(t)

        sm_result = None
        prso_result = None
        date = None

        for text in texts:
            lines = text.split("\n")

            for line in lines:

                raw_line = line.strip()
                line_clean = normalize(raw_line)

                # SM
                if not sm_result:
                    m = re.search(r"(SM\d{4}\.\d{4})", line_clean)
                    if m:
                        sm_result = m.group(1)

                # PR/SO (linh hoạt)
                if not prso_result:
                    m = re.search(r"(PR\d{4}\.\d{4}/?SO\d{4}\.\d{4})", line_clean)
                    if m:
                        prso_result = m.group(1)

                # DATE
                if not date:
                    d = re.search(r"(\d{2}/\d{2}/\d{4})", raw_line)
                    if d:
                        date = d.group(1)

        return sm_result, prso_result, date

    w, h = img.size

    variants = [
        img.crop((0,0,w,int(h*0.4))),
        img.crop((0,0,w,int(h*0.5))),
        img,
        img.rotate(180, expand=True).crop((0,0,w,int(h*0.4))),
        img.rotate(90, expand=True),
        img.rotate(270, expand=True),
    ]

    for variant in variants:
        sm, prso, date = read_lines(variant)

        if (sm or prso) and date:
            return sm, prso, date

    return None, None, None

# =========================
# PROCESS
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages_all):

    results = []
    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        processed_pages[0] += 1

        percent = int((i/total_pages)*100)
        global_percent = int((processed_pages[0] / total_pages_all) * 100)

        global_box.markdown(f"""
        <div class="global-bar">
            <div class="global-fill" style="width:{global_percent}%"></div>
        </div>
        """, unsafe_allow_html=True)

        box.markdown(f"""
<div class="file-row">
📄 {file.name} — Trang {i}/{total_pages} ({percent}%)
<div class="progress">
<div class="progress-bar" style="width:{percent}%"></div>
</div>
</div>
""", unsafe_allow_html=True)

        sm, prso, date = ocr_extract(img)

        if date:
            results.append({
                "SM": sm if sm else "",
                "PR/SO": prso if prso else "",
                "Ngày": date,
                "Trang": i
            })

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start_time = time.time()

        total_pages_all = sum(len(convert_from_bytes(f.read(), dpi=50)) for f in uploaded_files)
        for f in uploaded_files:
            f.seek(0)

        processed_pages = [0]

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for i, f in enumerate(uploaded_files):

                data = extract_pdf(
                    f, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        wb = load_workbook(tmp_excel.name)

        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border

            for cell in ws[1]:
                cell.font = Font(bold=True)

        wb.save(tmp_excel.name)

        st.session_state.excel_file = tmp_excel.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    with open(st.session_state.excel_file, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
        <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
