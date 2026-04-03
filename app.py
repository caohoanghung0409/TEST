import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import zipfile
import os
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="OCR Drive UI", layout="wide")

# =========================
# SESSION
# =========================
if "files_store" not in st.session_state:
    st.session_state.files_store = []

if "processing" not in st.session_state:
    st.session_state.processing = False

if "done" not in st.session_state:
    st.session_state.done = False

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.stApp { background: #f8fafc; }

.header {
    padding:15px;
    font-size:20px;
    font-weight:600;
}

/* UPLOADER */
[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5f5;
    padding: 40px;
    border-radius: 16px;
    text-align: center;
    background: white;
    cursor: pointer;
}

/* ẨN LIST MẶC ĐỊNH */
[data-testid="stFileUploader"] ul {
    display: none;
}

[data-testid="stFileUploader"] small { display: none; }
[data-testid="stFileUploader"] label { display: none; }

[data-testid="stFileUploader"]::before {
    content: "📤 Drag & Drop hoặc click để chọn PDF";
    display: block;
    font-size: 16px;
    color: #334155;
}

/* FILE LIST (GIỐNG uploader) */
.file-inline {
    display:flex;
    align-items:center;
    justify-content:space-between;
    background:#f1f5f9;
    padding:8px 12px;
    border-radius:8px;
    margin-top:6px;
    font-size:14px;
}

/* DELETE */
.del-btn {
    color:red;
    font-weight:bold;
    cursor:pointer;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">📁 OCR Drive Tool</div>', unsafe_allow_html=True)

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader("", type=["pdf"], accept_multiple_files=True)

# lưu file
if uploaded_files:
    for f in uploaded_files:
        if f.name not in [x["name"] for x in st.session_state.files_store]:
            st.session_state.files_store.append({
                "name": f.name,
                "file": f
            })

# =========================
# FILE LIST NGAY DƯỚI UPLOADER
# =========================
for i, f in enumerate(st.session_state.files_store):

    col1, col2 = st.columns([20,1])

    with col1:
        st.markdown(f"""
        <div class="file-inline">
            📄 {f["name"]}
        </div>
        """, unsafe_allow_html=True)

    with col2:
        if st.button("❌", key=f"del_{i}"):
            st.session_state.files_store.pop(i)
            st.rerun()

# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# PROCESS
# =========================
def extract_pdf(file, box, idx, total, global_bar):
    results = []
    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):
        percent = int((i/total_pages)*100)
        global_percent = int(((idx + i/total_pages)/total)*100)

        box.progress(percent)
        global_bar.progress(global_percent)

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)
        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    return results

# =========================
# MAIN
# =========================
if st.session_state.files_store:

    global_bar = st.progress(0)
    boxes = [st.empty() for _ in st.session_state.files_store]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Process Files"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            for i, f in enumerate(st.session_state.files_store):

                data = extract_pdf(f["file"], boxes[i], i, len(st.session_state.files_store), global_bar)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    name = os.path.splitext(f["name"])[0] + ".xlsx"

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

    with open(st.session_state.zip, "rb") as f:
        zip_data = f.read()

    if st.download_button(
        "📥 Download ZIP",
        zip_data,
        file_name="ocr_results.zip",
        mime="application/zip"
    ):
        st.markdown('<meta http-equiv="refresh" content="2">', unsafe_allow_html=True)
