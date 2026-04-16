import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import os
import time
import base64
from openpyxl import load_workbook
from io import BytesIO

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

if "excel_data" not in st.session_state:
    st.session_state.excel_data = None

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header {font-size:22px;font-weight:700;margin-bottom:10px;}

.global-bar {
    height:18px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
    margin-bottom:10px;
}
.global-fill {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}

.file-row {margin-top:12px;padding:10px;border-radius:12px;background:white;}

.progress {
    height:8px;background:#e5e7eb;border-radius:999px;overflow:hidden;margin-top:6px;
}
.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 THL PDF → EXCEL</div>', unsafe_allow_html=True)

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
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages):
    results = []
    images = convert_from_bytes(file.read(), dpi=150)

    for i, img in enumerate(images, start=1):
        processed_pages[0] += 1

        percent_file = int((i/len(images))*100)
        percent_global = int((processed_pages[0]/total_pages)*100)

        elapsed = time.time() - start_time
        speed = processed_pages[0] / elapsed if elapsed else 0
        eta = int((total_pages - processed_pages[0]) / speed) if speed else 0

        # GLOBAL BAR
        global_box.markdown(f"""
<div>⚡ {percent_global}% | 🚀 {speed:.2f} trang/s | ⏳ {eta}s</div>
<div class="global-bar">
<div class="global-fill" style="width:{percent_global}%"></div>
</div>
""", unsafe_allow_html=True)

        # FILE BAR
        box.markdown(f"""
<div class="file-row">
📄 {file.name} — Trang {i}/{len(images)} ({percent_file}%)
<div class="progress">
<div class="progress-bar" style="width:{percent_file}%"></div>
</div>
</div>
""", unsafe_allow_html=True)

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)
        if sm and date:
            results.append({"Trang": i, "SM": sm, "Ngày": date})

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

        total_pages = sum(len(convert_from_bytes(f.read(), dpi=50)) for f in uploaded_files)
        for f in uploaded_files:
            f.seek(0)

        processed_pages = [0]
        all_sheets = {}

        for i, f in enumerate(uploaded_files):

            data = extract_pdf(
                f, boxes[i], global_box,
                start_time, processed_pages, total_pages
            )

            if data:
                df = pd.DataFrame(data)
                df.insert(0, "STT", range(1, len(df)+1))
                all_sheets[os.path.splitext(f.name)[0][:31]] = df

        # EXPORT EXCEL
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)

        output.seek(0)

        # AUTO WIDTH
        wb = load_workbook(output)
        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.session_state.excel_data = final_output
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD + AUTO
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    data = st.session_state.excel_data.getvalue()
    b64 = base64.b64encode(data).decode()

    # AUTO DOWNLOAD
    st.markdown(f"""
        <iframe src="data:application/octet-stream;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    st.download_button(
        label="📥 Tải lại file",
        data=data,
        file_name="ket_qua.xlsx"
    )

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.processing = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
