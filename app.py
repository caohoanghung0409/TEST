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
st.set_page_config(page_title="OCR Timeline UI", layout="wide")

# =========================
# SESSION
# =========================
for key in ["processing", "done", "clear_uploader"]:
    if key not in st.session_state:
        st.session_state[key] = False

# =========================
# STYLE (TIMELINE UI)
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.stApp { background: #0f172a; color: #e2e8f0; }

.block-container {
    padding-top: 1rem !important;
}

/* PANEL */
.panel {
    background: #111827;
    padding: 20px;
    border-radius: 16px;
}

/* FILE BLOCK */
.file-block {
    border-left: 2px solid #334155;
    padding-left: 15px;
    margin-bottom: 20px;
}

/* DOT */
.dot {
    height: 10px;
    width: 10px;
    border-radius: 50%;
    display: inline-block;
    margin-right: 8px;
}

.running { background: #3b82f6; }
.done { background: #22c55e; }

/* TEXT */
.file-name {
    font-weight: 600;
}

.log {
    font-size: 13px;
    color: #94a3b8;
    margin-left: 18px;
    margin-top: 3px;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("## ⚡ OCR Processing Timeline")

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "Upload PDF",
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
# MAIN
# =========================
if uploaded_files:

    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Start Processing"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            for idx, f in enumerate(uploaded_files):

                images = convert_from_bytes(f.read(), dpi=150)
                total_pages = len(images)
                results = []

                for i, img in enumerate(images, start=1):

                    boxes[idx].markdown(f"""
<div class="panel">
    <div class="file-block">
        <div class="file-name">
            <span class="dot running"></span>{f.name}
        </div>
        <div class="log">Processing page {i}/{total_pages}</div>
        <div class="log">Running OCR...</div>
    </div>
</div>
""", unsafe_allow_html=True)

                    w, h = img.size
                    img = img.crop((0, 0, w, int(h * 0.4)))

                    sm, date = process_page(img)
                    if sm and date:
                        results.append({"SM": sm, "Ngày": date})

                # DONE UI
                boxes[idx].markdown(f"""
<div class="panel">
    <div class="file-block">
        <div class="file-name">
            <span class="dot done"></span>{f.name}
        </div>
        <div class="log">Completed {total_pages} pages</div>
        <div class="log">Data extracted successfully</div>
    </div>
</div>
""", unsafe_allow_html=True)

                # EXPORT
                if results:
                    df = pd.DataFrame(results)
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

    st.success("🎉 All files processed!")

    with open(st.session_state.zip, "rb") as f:
        data = f.read()

    if st.download_button("📥 Download ZIP", data, file_name="ocr_results.zip"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
