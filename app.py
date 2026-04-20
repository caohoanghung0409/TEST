import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
import pandas as pd
import re
import time
import numpy as np
import cv2
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF → EXCEL", layout="wide")

# =========================
# SESSION
# =========================
if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None
if "file_name" not in st.session_state:
    st.session_state.file_name = None
if "result" not in st.session_state:
    st.session_state.result = None

# =========================
# UI HEADER
# =========================
st.markdown("## 🚀 THL PDF → EXCEL")

files = st.file_uploader("📂 Chọn PDF", type=["pdf"], accept_multiple_files=False)

# =========================
# OCR IMPROVED
# =========================
SM_REGEX = re.compile(r"(SM\d{4}\.\d{4})")
DATE_REGEX = re.compile(r"(\d{2}/\d{2}/\d{4})")


def preprocess(img):
    """🔥 tăng OCR accuracy"""
    img = np.array(img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return gray


def ocr(img):

    img = preprocess(img)

    text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return sm.group(1), date.group(1)

    # fallback rotate
    img_rot = cv2.rotate(img, cv2.ROTATE_180)
    text2 = pytesseract.image_to_string(img_rot, config="--oem 3 --psm 6")

    sm = SM_REGEX.search(text2)
    date = DATE_REGEX.search(text2)

    if sm and date:
        return sm.group(1), date.group(1)

    return None, None


# =========================
# PROCESS
# =========================
def process_pdf(file_bytes, file_name, progress_box):

    pages = convert_from_bytes(file_bytes, dpi=150)
    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"

    ws.append(["STT", "SM", "Ngày", "Trang"])

    total = len(pages)

    for i, img in enumerate(pages, start=1):

        percent = int(i / total * 100)

        progress_box.markdown(f"⏳ {percent}% ({i}/{total})")

        sm, date = ocr(img)

        if sm and date:
            ws.append([i, sm, date, i])

    # format
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for c in row:
            c.border = border
            if row[0].row == 1:
                c.font = Font(bold=True)

    return wb


# =========================
# RUN
# =========================
if files and st.button("🚀 BẮT ĐẦU"):

    progress_box = st.empty()

    file_bytes = files.read()

    file_name = files.name.replace(".pdf", ".xlsx")

    with st.spinner("Đang xử lý OCR..."):

        wb = process_pdf(file_bytes, files.name, progress_box)

        # save to buffer (KHÔNG dùng temp file)
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.session_state.result = buffer
        st.session_state.file_name = file_name

# =========================
# DOWNLOAD (CHUẨN STREAMLIT)
# =========================
if st.session_state.result:

    st.success("🎉 HOÀN THÀNH !!!")

    st.download_button(
        label="📥 TẢI FILE EXCEL",
        data=st.session_state.result,
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
