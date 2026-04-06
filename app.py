import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import zipfile
import os
import time
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="OCR Drive UI", layout="wide")

# =========================
# SESSION
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False

if "done" not in st.session_state:
    st.session_state.done = False

if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

# =========================
# STYLE MAX PRO (FIX UI)
# =========================
st.markdown("""
<style>
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* 🔥 KÉO NỘI DUNG LÊN TRÊN */
.block-container {
    padding-top: 0.5rem !important;
}

/* nền */
.stApp { background: #f8fafc; }

/* header */
.header {
    padding:5px 0;
    font-size:20px;
    font-weight:600;
}

/* uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5f5;
    padding: 20px;
    border-radius: 16px;
    background: white;
}

/* file */
.file-row {
    font-size:14px;
    margin-top:10px;
}

.file-name {
    font-weight:500;
}

.file-status {
    color:#64748b;
    font-size:13px;
}

/* progress file */
.progress {
    height:6px;
    background:#e5e7eb;
    border-radius:10px;
    overflow:hidden;
    margin-top:4px;
}

.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#0ea5e9,#22c55e);
}

/* global */
.global-wrap {
    margin:10px 0;
}

.global-bar {
    position:relative;
    height:18px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
}

.global-fill {
    height:100%;
    border-radius:999px;
    transition: width 0.4s ease;
}

.global-fill::before {
    content:"";
    position:absolute;
    width:100%;
    height:100%;
    background: repeating-linear-gradient(
        45deg,
        rgba(255,255,255,0.2) 0,
        rgba(255,255,255,0.2) 10px,
        transparent 10px,
        transparent 20px
    );
    animation: move 1s linear infinite;
}

@keyframes move {
    from { background-position: 0 0; }
    to { background-position: 40px 0; }
}

.global-text {
    position:absolute;
    width:100%;
    text-align:center;
    font-size:12px;
    font-weight:600;
    top:0;
    line-height:18px;
    color:#0f172a;
}

.global-meta {
    display:flex;
    justify-content:space-between;
    font-size:12px;
    margin-bottom:5px;
    color:#475569;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">📁 CHECK PDF TO EXCEL ( SM ) </div>', unsafe_allow_html=True)

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# COLOR
# =========================
def get_color(percent):
    if percent < 30:
        return "#0ea5e9"
    elif percent < 70:
        return "#f59e0b"
    elif percent < 100:
        return "#ef4444"
    else:
        return "#22c55e"

# =========================
# GLOBAL BAR
# =========================
def render_global_bar(percent, speed, eta):
    color = get_color(percent)

    return f"""
<div class="global-wrap">
    <div class="global-meta">
        <div>⚡ {percent}%</div>
        <div>🚀 {speed:.2f} pages/s • ⏳ ETA: {eta}s</div>
    </div>
    <div class="global-bar">
        <div class="global-fill" style="width:{percent}%; background:{color};"></div>
        <div class="global-text">{percent}%</div>
    </div>
</div>
"""

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages_all):
    results = []
    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):
        processed_pages[0] += 1

        percent = int((i/total_pages)*100)
        global_percent = int((processed_pages[0] / total_pages_all) * 100)

        elapsed = time.time() - start_time
        speed = processed_pages[0] / elapsed if elapsed > 0 else 0
        remaining = total_pages_all - processed_pages[0]
        eta = int(remaining / speed) if speed > 0 else 0

        global_box.markdown(
            render_global_bar(global_percent, speed, eta),
            unsafe_allow_html=True
        )

        box.markdown(f"""
<div class="file-row">
    <div class="file-name">📄 {file.name}</div>
    <div class="file-status">Trang {i}/{total_pages} • {percent}%</div>
    <div class="progress">
        <div class="progress-bar" style="width:{percent}%"></div>
    </div>
</div>
""", unsafe_allow_html=True)

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)
        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Process Files"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start_time = time.time()

        total_pages_all = sum(len(convert_from_bytes(f.read(), dpi=50)) for f in uploaded_files)
        for f in uploaded_files:
            f.seek(0)

        processed_pages = [0]

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            for i, f in enumerate(uploaded_files):

                data = extract_pdf(
                    f, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    name = os.path.splitext(f.name)[0] + ".xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name, index=False)

                        wb = load_workbook(tmp.name)
                        ws = wb.active

                        for col in ws.columns:
                            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                            ws.column_dimensions[col[0].column_letter].width = max_len + 3

                        wb.save(tmp.name)
                        zipf.write(tmp.name, name)

        st.session_state.zip = zip_buffer.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 Xử lý xong!")

    with open(st.session_state.zip, "rb") as f:
        zip_data = f.read()

    if st.download_button(
        "📥 Download ZIP",
        zip_data,
        file_name="ocr_results.zip",
        mime="application/zip"
    ):
        st.toast("✅ Download xong!", icon="🎉")

        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
