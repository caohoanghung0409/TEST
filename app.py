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
# UI STYLE (FINAL)
# =========================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #0ea5e9, #22c55e);
}

/* Title */
h1 {
    text-align: center;
    color: white !important;
    font-weight: 900;
}

/* Card */
.card {
    background: rgba(255,255,255,0.95);
    padding: 16px;
    border-radius: 16px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    margin-bottom: 10px;
}

/* File name */
.file-name {
    font-weight: 700;
    font-size: 16px;
}

/* Status */
.status {
    font-size: 14px;
    margin-top: 6px;
}

/* Progress bar */
.progress-bar {
    height: 8px;
    border-radius: 10px;
    background: #e5e7eb;
    margin-top: 10px;
    overflow: hidden;
}

.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #0284c7, #22c55e);
    transition: width 0.3s ease-in-out;
}

/* Button */
.stButton > button {
    width: 100%;
    background: linear-gradient(90deg, #0284c7, #22c55e);
    color: white;
    border-radius: 12px;
    padding: 12px;
    font-weight: 700;
    border: none;
}
</style>
""", unsafe_allow_html=True)

st.title("📄 OCR MULTI PDF → EXCEL (HYBRID UI)")

# =========================
# OCR FUNCTION
# =========================
def process_page(img):
    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )

    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    return (sm.group(1), date.group(1)) if sm and date else (None, None)


# =========================
# UPDATE CARD (CLEAN UI)
# =========================
def update_card(box, file_name, page, total):

    percent = round((page / total) * 100)

    box.markdown(f"""
    <div class="card">
        <div class="file-name">📁 {file_name}</div>

        <div class="status">
            📄 {page} / {total} &nbsp;&nbsp; • &nbsp;&nbsp;
            ⚡ Processing &nbsp;&nbsp; • &nbsp;&nbsp;
            {percent}%
        </div>

        <div class="progress-bar">
            <div class="progress-fill" style="width:{percent}%"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# =========================
# DONE CARD
# =========================
def update_done(box, file_name):

    box.markdown(f"""
    <div class="card">
        <div class="file-name">📁 {file_name}</div>

        <div class="status">
            ✅ Hoàn tất
        </div>

        <div class="progress-bar">
            <div class="progress-fill" style="width:100%"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# =========================
# PROCESS PDF
# =========================
def extract_pdf(file, box):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        update_card(box, file.name, i, total_pages)

        # crop tăng tốc
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date
            })

    update_done(box, file.name)

    return results


# =========================
# AUTO WIDTH
# =========================
def auto_width(path):
    wb = load_workbook(path)
    ws = wb.active

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_len + 3

    wb.save(path)


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

    st.success(f"📦 Đã chọn {len(uploaded_files)} file")

    # responsive grid (max 4 cột)
    cols = st.columns(min(len(uploaded_files), 4))

    boxes = []
    for i, file in enumerate(uploaded_files):
        with cols[i % 4]:
            box = st.empty()
            boxes.append(box)

    if st.button("🚀 START OCR"):

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for idx, file in enumerate(uploaded_files):

                data = extract_pdf(file, boxes[idx])

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))

                    base = os.path.splitext(file.name)[0]
                    excel_name = f"{base}.xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name, index=False)
                        auto_width(tmp.name)
                        zipf.write(tmp.name, excel_name)

        st.success("🎉 HOÀN TẤT TẤT CẢ!")

        with open(zip_buffer.name, "rb") as f:
            st.download_button(
                "📥 DOWNLOAD ZIP",
                f,
                file_name="ocr_results.zip"
            )
