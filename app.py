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
# SAAS UI STYLE
# =========================
st.markdown("""
<style>

/* HIDE STREAMLIT */
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* BACKGROUND */
.stApp {
    background: linear-gradient(135deg, #0f172a, #1e293b);
    color: white;
}

/* HEADER */
.header {
    padding: 20px;
    border-radius: 16px;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    box-shadow: 0 8px 30px rgba(0,0,0,0.3);
    text-align: center;
    margin-bottom: 20px;
}

.header h1 {
    margin: 0;
    font-size: 32px;
}

.header p {
    margin: 0;
    opacity: 0.9;
}

/* STATS */
.stats {
    display: flex;
    gap: 15px;
    margin-bottom: 20px;
}

.stat-card {
    flex: 1;
    background: rgba(255,255,255,0.08);
    padding: 15px;
    border-radius: 14px;
    text-align: center;
    backdrop-filter: blur(10px);
}

/* CARD */
.card {
    background: rgba(255,255,255,0.08);
    padding: 15px;
    border-radius: 16px;
    backdrop-filter: blur(10px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.3);
}

/* FILE NAME */
.file-name {
    font-weight: 700;
    font-size: 14px;
    margin-bottom: 5px;
}

/* PROGRESS */
.progress-bar {
    width: 100%;
    height: 8px;
    background: rgba(255,255,255,0.2);
    border-radius: 10px;
    overflow: hidden;
    margin-top: 10px;
}

.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #22c55e, #4ade80);
    transition: width 0.3s;
}

/* BUTTON */
.stButton>button {
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
    font-weight: 700;
    border-radius: 12px;
    padding: 10px 20px;
    border: none;
}

</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>📄 OCR PDF SaaS Dashboard</h1>
    <p>Extract SM & Date from multiple PDFs automatically</p>
</div>
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
def extract_pdf(file, status_box):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        percent = int((i / total_pages) * 100)

        html = f"""
<div class="card">
    <div class="file-name">📁 {file.name}</div>
    <div>📄 {i}/{total_pages} • Processing ({percent}%)</div>
    <div class="progress-bar">
        <div class="progress-fill" style="width:{percent}%"></div>
    </div>
</div>
""".strip()

        status_box.markdown(html, unsafe_allow_html=True)

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    status_box.markdown(f"""
<div class="card">
    <div class="file-name">📁 {file.name}</div>
    <div style="color:#4ade80;font-weight:700;">✅ DONE</div>
</div>
""", unsafe_allow_html=True)

    return results


# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📤 Upload PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# STATS BAR
# =========================
if uploaded_files:

    st.markdown(f"""
    <div class="stats">
        <div class="stat-card">📦 Files<br><b>{len(uploaded_files)}</b></div>
        <div class="stat-card">⚡ Status<br><b>Ready</b></div>
    </div>
    """, unsafe_allow_html=True)

    cols = st.columns(len(uploaded_files))
    status_boxes = []

    for i in range(len(uploaded_files)):
        with cols[i]:
            status_boxes.append(st.empty())

    if st.button("🚀 Start Processing"):

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for idx, file in enumerate(uploaded_files):

                data = extract_pdf(file, status_boxes[idx])

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

        st.success("🎉 All files processed!")

        with open(zip_buffer.name, "rb") as f:
            st.download_button("📥 Download ZIP", f, file_name="ocr_results.zip")
