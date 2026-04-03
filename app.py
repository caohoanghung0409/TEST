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
st.set_page_config(page_title="OCR SaaS PRO", layout="wide")

# =========================
# SESSION STATE
# =========================
if "done" not in st.session_state:
    st.session_state.done = False

if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.title("⚙️ OCR Tool")
    theme = st.radio("🎨 Theme", ["Light", "Dark"])

# =========================
# THEME
# =========================
if theme == "Dark":
    bg = "#0f172a"
    card = "rgba(255,255,255,0.08)"
    text = "white"
else:
    bg = "#f1f5f9"
    card = "white"
    text = "#0f172a"

st.markdown(f"""
<style>
header {{visibility: hidden;}}
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

.stApp {{
    background: {bg};
    color: {text};
}}

.card {{
    background: {card};
    padding: 15px;
    border-radius: 14px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.08);
}}

.progress-bar {{
    width: 100%;
    height: 8px;
    background: #e5e7eb;
    border-radius: 10px;
    overflow: hidden;
    margin-top: 10px;
}}

.progress-fill {{
    height: 100%;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
}}

.stButton>button {{
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
    font-weight: 700;
    border-radius: 10px;
}}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("""
<h2 style='text-align:center;'>📄 OCR PDF Dashboard</h2>
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
# PROCESS FILE
# =========================
def extract_pdf(file, status_box, global_progress, idx, total_files):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        percent = int((i / total_pages) * 100)
        global_percent = int(((idx + (i/total_pages)) / total_files) * 100)

        html = f"""
<div class="card">
📁 {file.name}<br>
📄 {i}/{total_pages} • {percent}%
<div class="progress-bar">
<div class="progress-fill" style="width:{percent}%"></div>
</div>
</div>
"""

        status_box.markdown(html, unsafe_allow_html=True)
        global_progress.progress(global_percent)

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    status_box.markdown(f"""
<div class="card">📁 {file.name}<br>✅ DONE</div>
""", unsafe_allow_html=True)

    return results

# =========================
# UPLOADER (RESET KEY)
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "📤 Upload PDF",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

# =========================
# MAIN
# =========================
if uploaded_files:

    global_progress = st.progress(0)

    cols = st.columns(len(uploaded_files))
    status_boxes = [cols[i].empty() for i in range(len(uploaded_files))]

    if st.button("🚀 Start OCR"):

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for idx, file in enumerate(uploaded_files):

                data = extract_pdf(
                    file,
                    status_boxes[idx],
                    global_progress,
                    idx,
                    len(uploaded_files)
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))

                    base = os.path.splitext(file.name)[0]
                    excel_name = f"{base}.xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name, index=False)

                        wb = load_workbook(tmp.name)
                        ws = wb.active

                        for col in ws.columns:
                            max_len = 0
                            col_letter = col[0].column_letter
                            for c in col:
                                if c.value:
                                    max_len = max(max_len, len(str(c.value)))
                            ws.column_dimensions[col_letter].width = max_len + 3

                        wb.save(tmp.name)
                        zipf.write(tmp.name, excel_name)

        st.session_state.done = True
        st.session_state.zip_path = zip_buffer.name

# =========================
# DOWNLOAD + SMOOTH RESET
# =========================
if st.session_state.done:

    st.success("🎉 Xử lý xong!")

    with open(st.session_state.zip_path, "rb") as f:
        if st.download_button("📥 Download ZIP", f, file_name="ocr_results.zip"):

            # 🔥 RESET MƯỢT (KHÔNG RELOAD)
            st.session_state.done = False
            st.session_state.clear_uploader = not st.session_state.clear_uploader

            # Toast kiểu SaaS
            st.toast("✅ Download thành công!", icon="🎉")

            st.rerun()
