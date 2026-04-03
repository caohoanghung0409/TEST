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
st.set_page_config(page_title="OCR PDF Tool", layout="wide")

# =========================
# UI STYLE (NEW MODERN)
# =========================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #0ea5e9, #22c55e);
}

/* TITLE */
h1 {
    text-align: center;
    color: white !important;
    font-weight: 900;
    letter-spacing: 1px;
}

/* CARD */
.card {
    background: rgba(255,255,255,0.95);
    padding: 15px;
    border-radius: 16px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.2);
    text-align: center;
    transition: 0.3s;
}

.card:hover {
    transform: translateY(-3px);
}

/* FILE NAME */
.file-name {
    font-weight: 700;
    color: #0284c7;
    margin-bottom: 8px;
}

/* STATUS TEXT */
.status {
    font-size: 14px;
    margin-top: 5px;
}

/* PROGRESS BAR */
.progress-bar {
    width: 100%;
    height: 10px;
    background: #e5e7eb;
    border-radius: 10px;
    overflow: hidden;
    margin-top: 10px;
}

.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #22c55e, #16a34a);
    transition: width 0.3s ease;
}

/* BUTTON */
.stButton>button {
    background: linear-gradient(90deg, #0284c7, #22c55e);
    color: white;
    font-weight: 700;
    border-radius: 10px;
    padding: 10px 20px;
    border: none;
}

.stButton>button:hover {
    transform: scale(1.05);
}

/* SUCCESS BOX */
.success-box {
    background: #ecfdf5;
    padding: 12px;
    border-radius: 10px;
    font-weight: 600;
    color: #16a34a;
}

</style>
""", unsafe_allow_html=True)

st.title("📄 OCR MULTI PDF DASHBOARD PRO")


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

        status_box.markdown(
            f"""
            <div class="card">
                <div class="file-name">📁 {file.name}</div>
                <div class="status">📄 {i}/{total_pages} | ⚡ Processing ({percent}%)</div>

                <div class="progress-bar">
                    <div class="progress-fill" style="width:{percent}%"></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    status_box.markdown(
        f"""
        <div class="card">
            <div class="file-name">📁 {file.name}</div>
            <div class="success-box">✅ DONE</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    return results


# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📤 Upload nhiều PDF",
    type=["pdf"],
    accept_multiple_files=True
)


# =========================
# RUN
# =========================
if uploaded_files:

    st.markdown(
        f"""
        <div class="card">
            📦 Tổng file: <b>{len(uploaded_files)}</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    cols = st.columns(len(uploaded_files))
    status_boxes = []

    for i in range(len(uploaded_files)):
        with cols[i]:
            box = st.empty()
            status_boxes.append(box)

    if st.button("🚀 START OCR"):

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

        st.success("🎉 HOÀN TẤT!")

        with open(zip_buffer.name, "rb") as f:
            st.download_button(
                "📥 DOWNLOAD ZIP",
                f,
                file_name="ocr_results.zip"
            )
