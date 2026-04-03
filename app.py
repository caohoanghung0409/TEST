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
# PAGE CONFIG
# =========================
st.set_page_config(page_title="OCR PDF Tool", layout="wide")

# =========================
# MODERN UI (Dashboard style)
# =========================
st.markdown("""
<style>

/* nền app */
.stApp {
    background: linear-gradient(135deg, #0ea5e9, #22c55e);
}

/* title */
h1 {
    text-align: center;
    color: white !important;
    font-size: 40px !important;
    font-weight: 900;
    margin-bottom: 10px;
}

/* container */
.block-container {
    padding: 1.5rem;
}

/* uploader */
div[data-testid="stFileUploader"] {
    background: white;
    padding: 15px;
    border-radius: 14px;
    border: 2px dashed #22c55e;
}

/* button */
.stButton > button {
    width: 100%;
    background: linear-gradient(90deg, #0284c7, #22c55e);
    color: white;
    border-radius: 12px;
    padding: 12px;
    font-weight: 700;
    border: none;
}

/* status box */
.status-box {
    background: rgba(255,255,255,0.92);
    padding: 12px 16px;
    border-radius: 12px;
    margin-top: 10px;
    font-weight: 600;
    box-shadow: 0 8px 20px rgba(0,0,0,0.12);
}

/* progress */
.stProgress > div > div {
    background: linear-gradient(90deg, #0284c7, #22c55e);
}

</style>
""", unsafe_allow_html=True)

st.title("📄 OCR MULTI PDF → EXCEL DASHBOARD")


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

    if sm and date:
        return sm.group(1), date.group(1)

    return None, None


# =========================
# EXTRACT PDF (ONE LINE STATUS)
# =========================
def extract_pdf(file, file_index, total_files, status_holder):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    page_progress = st.progress(0)

    for i, img in enumerate(images, start=1):

        # ✅ GOM 1 DÒNG STATUS
        status_holder.markdown(
            f"""
            <div class="status-box">
            📁 File {file_index}/{total_files} | {file.name}  
            📄 Trang {i}/{total_pages} | ⚡ Đang xử lý...
            </div>
            """,
            unsafe_allow_html=True
        )

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({"SM": sm, "Ngày": date})

        page_progress.progress(i / total_pages)

    return results


# =========================
# AUTO WIDTH EXCEL
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
# PROCESS ALL
# =========================
if uploaded_files:

    st.success(f"📦 Đã chọn {len(uploaded_files)} file")

    if st.button("🚀 START PROCESS"):
        main_progress = st.progress(0)
        main_status = st.empty()

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
        total_files = len(uploaded_files)

        status_holder = st.empty()

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for idx, file in enumerate(uploaded_files, start=1):

                main_status.info(f"⚡ Processing {file.name} ({idx}/{total_files})")

                data = extract_pdf(file, idx, total_files, status_holder)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))

                    base_name = os.path.splitext(file.name)[0]
                    excel_name = f"{base_name}.xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name, index=False)
                        auto_width(tmp.name)

                        zipf.write(tmp.name, excel_name)

                main_progress.progress(idx / total_files)

        main_status.success("🎉 HOÀN TẤT TẤT CẢ FILE!")

        with open(zip_buffer.name, "rb") as f:
            st.download_button(
                "📥 DOWNLOAD ALL (ZIP)",
                f,
                file_name="ocr_results.zip"
            )
