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
# UI
# =========================
st.set_page_config(page_title="OCR PDF Tool", layout="wide")
st.title("📄 OCR Nhiều PDF → Excel (Hiển thị từng trang)")

# =========================
# OCR PAGE
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
# EXTRACT 1 FILE (WITH PAGE PROGRESS)
# =========================
def extract_pdf(file, file_index, total_files):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    file_status = st.empty()
    page_progress = st.progress(0)

    for i, img in enumerate(images, start=1):

        # 👉 hiển thị rõ file + trang
        file_status.text(
            f"📁 File {file_index}/{total_files}: {file.name} | 📄 Trang {i}/{total_pages}"
        )

        # crop tăng tốc
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date
            })

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
# UI INPUT
# =========================
uploaded_files = st.file_uploader(
    "📤 Upload nhiều PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# PROCESS ALL FILES
# =========================
if uploaded_files:
    if st.button("🚀 Xử lý tất cả"):
        main_progress = st.progress(0)
        main_status = st.empty()

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        total_files = len(uploaded_files)

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for idx, file in enumerate(uploaded_files, start=1):

                main_status.text(f"⚡ Đang xử lý file {idx}/{total_files}: {file.name}")

                data = extract_pdf(file, idx, total_files)

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

        main_status.text("✅ Hoàn tất tất cả file!")

        with open(zip_buffer.name, "rb") as f:
            st.download_button(
                "📥 Tải tất cả (ZIP)",
                f,
                file_name="ocr_results.zip"
            )
