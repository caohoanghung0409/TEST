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
if "done" not in st.session_state:
    st.session_state.done = False

if "clear" not in st.session_state:
    st.session_state.clear = False

# =========================
# STYLE (GOOGLE DRIVE)
# =========================
st.markdown("""
<style>
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.stApp {
    background: #f8fafc;
}

/* HEADER */
.header {
    display:flex;
    align-items:center;
    gap:10px;
    padding:15px;
    font-size:20px;
    font-weight:600;
}

/* UPLOAD BOX */
.upload-box {
    border: 2px dashed #cbd5f5;
    padding: 40px;
    border-radius: 16px;
    text-align:center;
    background:white;
    margin-bottom:20px;
}

/* FILE LIST */
.file-row {
    background:white;
    padding:15px;
    border-radius:12px;
    margin-bottom:10px;
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}

/* PROGRESS */
.progress {
    height:6px;
    background:#e5e7eb;
    border-radius:10px;
    overflow:hidden;
    margin-top:8px;
}

.progress-bar {
    height:100%;
    background:#0ea5e9;
}

/* BUTTON */
.stButton>button {
    background:#0ea5e9;
    color:white;
    border-radius:8px;
    font-weight:600;
}

</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("""
<div class="header">📁 OCR Drive Tool</div>
""", unsafe_allow_html=True)

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

        html = f"""
<div class="file-row">
📄 {file.name}<br>
{i}/{total_pages} • {percent}%
<div class="progress">
<div class="progress-bar" style="width:{percent}%"></div>
</div>
</div>
"""
        box.markdown(html, unsafe_allow_html=True)
        global_bar.progress(global_percent)

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    box.markdown(f"""
<div class="file-row">📄 {file.name} ✅ DONE</div>
""", unsafe_allow_html=True)

    return results

# =========================
# UPLOAD BOX
# =========================
st.markdown('<div class="upload-box">📤 Drag & Drop PDF hoặc chọn file bên dưới</div>', unsafe_allow_html=True)

uploader_key = "u1" if not st.session_state.clear else "u2"

files = st.file_uploader(
    "",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

# =========================
# MAIN
# =========================
if files:

    st.write(f"📦 {len(files)} file đã chọn")

    global_bar = st.progress(0)

    boxes = [st.empty() for _ in files]

    if st.button("🚀 Process Files"):

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for i, f in enumerate(files):

                data = extract_pdf(f, boxes[i], i, len(files), global_bar)

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

        st.session_state.done = True
        st.session_state.zip = zip_buffer.name

# =========================
# DOWNLOAD + RESET
# =========================
if st.session_state.done:

    st.success("🎉 Xong!")

    with open(st.session_state.zip, "rb") as f:
        if st.download_button("📥 Download ZIP", f):

            st.toast("✅ Download xong!", icon="🎉")

            # reset mượt
            st.session_state.done = False
            st.session_state.clear = not st.session_state.clear

            st.rerun()
