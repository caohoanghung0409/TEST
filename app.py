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
st.set_page_config(page_title="OCR SaaS", layout="wide")

# =========================
# SESSION
# =========================
for key in ["processing", "done", "clear_uploader"]:
    if key not in st.session_state:
        st.session_state[key] = False

# =========================
# STYLE (SaaS UI)
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.stApp { background: #f1f5f9; }

.block-container {
    padding-top: 1rem !important;
}

/* LAYOUT */
.container {
    display: flex;
    gap: 20px;
}

/* PANEL */
.panel {
    background: white;
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.05);
}

/* TITLE */
.title {
    font-size: 20px;
    font-weight: 600;
    margin-bottom: 10px;
}

/* FILE ITEM */
.file-item {
    padding: 10px 0;
    border-bottom: 1px solid #e2e8f0;
}

.file-name {
    font-weight: 500;
}

.status {
    font-size: 13px;
    color: #64748b;
}

/* PROGRESS */
.progress {
    height: 6px;
    background: #e2e8f0;
    border-radius: 999px;
    overflow: hidden;
    margin-top: 5px;
    position: relative;
}

.bar {
    height: 100%;
    background: linear-gradient(90deg,#3b82f6,#22c55e);
    transition: width 0.3s ease;
}

/* SHIMMER */
.progress::before {
    content: "";
    position: absolute;
    width: 30%;
    height: 100%;
    left: -30%;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.6), transparent);
    animation: shimmer 1.2s infinite;
}

@keyframes shimmer {
    100% { left: 130%; }
}

/* DONE STATE */
.done .progress::before {
    display: none;
}

/* BUTTON */
button[kind="primary"] {
    background: linear-gradient(90deg,#3b82f6,#6366f1);
    border: none;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("### 🚀 OCR SaaS Tool")

# =========================
# LAYOUT
# =========================
col1, col2 = st.columns([1,2])

# =========================
# LEFT PANEL (UPLOAD)
# =========================
with col1:
    st.markdown('<div class="panel">', unsafe_allow_html=True)

    uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

    uploaded_files = st.file_uploader(
        "Upload PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=uploader_key
    )

    if uploaded_files and not st.session_state.processing:
        if st.button("🚀 Process"):
            st.session_state.processing = True
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# =========================
# OCR FUNC
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# RIGHT PANEL (STATUS)
# =========================
with col2:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="title">Processing Status</div>', unsafe_allow_html=True)

    if uploaded_files:

        global_bar = st.progress(0)
        boxes = [st.empty() for _ in uploaded_files]

        if st.session_state.processing:

            zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

            with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
                for idx, f in enumerate(uploaded_files):

                    images = convert_from_bytes(f.read(), dpi=150)
                    total_pages = len(images)
                    results = []

                    for i, img in enumerate(images, start=1):
                        percent = int((i/total_pages)*100)
                        global_percent = int(((idx + i/total_pages)/len(uploaded_files))*100)

                        boxes[idx].markdown(f"""
<div class="file-item">
    <div class="file-name">📄 {f.name}</div>
    <div class="status">Processing • Page {i}/{total_pages} • {percent}%</div>
    <div class="progress">
        <div class="bar" style="width:{percent}%"></div>
    </div>
</div>
""", unsafe_allow_html=True)

                        global_bar.progress(global_percent)

                        w,h = img.size
                        img = img.crop((0,0,w,int(h*0.4)))

                        sm, date = process_page(img)
                        if sm and date:
                            results.append({"SM": sm, "Ngày": date})

                    # DONE
                    boxes[idx].markdown(f"""
<div class="file-item done">
    <div class="file-name">📄 {f.name}</div>
    <div class="status">✅ Done</div>
    <div class="progress">
        <div class="bar" style="width:100%"></div>
    </div>
</div>
""", unsafe_allow_html=True)

                    # EXPORT
                    if results:
                        df = pd.DataFrame(results)
                        df.insert(0, "STT", range(1,len(df)+1))

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

    st.markdown('</div>', unsafe_allow_html=True)

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:
    st.success("🎉 Done!")

    with open(st.session_state.zip, "rb") as f:
        data = f.read()

    if st.download_button("📥 Download ZIP", data, file_name="ocr.zip"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
