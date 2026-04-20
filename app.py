import streamlit as st
import fitz  # PyMuPDF
import re
import numpy as np
import cv2
from paddleocr import PaddleOCR
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font
from io import BytesIO

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="FAST PDF OCR", layout="wide")
st.title("🚀 PDF → EXCEL (PyMuPDF + PaddleOCR)")

# =========================
# OCR ENGINE (LOAD 1 LẦN)
# =========================
ocr = PaddleOCR(use_angle_cls=True, lang='en')

# =========================
# REGEX
# =========================
SM_REGEX = re.compile(r"SM\s*[-:]?\s*\d{4}\s*\.?\s*\d{4}")
DATE_REGEX = re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}")

# =========================
# IMAGE PREPROCESS (NHẸ + NHANH)
# =========================
def preprocess(img):

    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    img = cv2.resize(img, None, fx=1.5, fy=1.5)

    img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY)[1]

    return img

# =========================
# OCR FUNCTION
# =========================
def run_ocr(img):

    img = preprocess(img)

    result = ocr.ocr(img, cls=True)

    text = ""
    if result and result[0]:
        for line in result[0]:
            text += line[1][0] + " "

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return sm.group(), date.group()

    return None, None

# =========================
# PROCESS PDF (FAST CORE)
# =========================
def process_pdf(file_bytes, progress_box):

    doc = fitz.open(stream=file_bytes, filetype="pdf")

    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"

    ws.append(["STT", "SM", "Ngày", "Trang"])

    total_pages = len(doc)
    found = 0

    for i in range(total_pages):

        page = doc.load_page(i)

        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # zoom 2x

        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

        sm, date = run_ocr(img)

        if sm and date:
            found += 1
            ws.append([found, sm, date, i + 1])

        percent = int((i + 1) / total_pages * 100)
        progress_box.markdown(f"⚡ {percent}% ({i+1}/{total_pages})")

    # =========================
    # FORMAT EXCEL
    # =========================
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            if row[0].row == 1:
                cell.font = Font(bold=True)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer, found

# =========================
# UI
# =========================
file = st.file_uploader("📂 Chọn PDF", type=["pdf"])

if file and st.button("🚀 START"):

    progress = st.empty()

    file_bytes = file.read()

    with st.spinner("Đang xử lý PyMuPDF + PaddleOCR..."):

        result, count = process_pdf(file_bytes, progress)

    # =========================
    # RESULT
    # =========================
    if count == 0:
        st.error("❌ Không đọc được dữ liệu (PDF quá mờ hoặc không đúng format)")
    else:
        st.success(f"🎉 DONE - {count} dòng dữ liệu")

        st.download_button(
            "📥 DOWNLOAD EXCEL",
            data=result,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
