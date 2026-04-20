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

st.title("🚀 THL PDF → EXCEL (OCR FIXED)")

# =========================
# REGEX (LINH HOẠT HƠN)
# =========================
SM_REGEX = re.compile(r"SM\s*[-:]?\s*\d{4}\s*\.?\s*\d{4}")
DATE_REGEX = re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}")

# =========================
# OCR (BẢN MẠNH - FIX TRỐNG DATA)
# =========================
def ocr(img):

    img = np.array(img)

    # 🔥 upscale giúp đọc rõ hơn
    img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 🔥 tăng chất lượng ảnh
    gray = cv2.bilateralFilter(gray, 9, 75, 75)
    gray = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        2
    )

    config = "--oem 3 --psm 6"

    text = pytesseract.image_to_string(gray, config=config)

    # 🔥 debug (bật nếu cần)
    # st.text(text)

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return sm.group(), date.group()

    # 🔥 fallback rotate
    rot = cv2.rotate(gray, cv2.ROTATE_180)
    text2 = pytesseract.image_to_string(rot, config=config)

    sm = SM_REGEX.search(text2)
    date = DATE_REGEX.search(text2)

    if sm and date:
        return sm.group(), date.group()

    return None, None


# =========================
# PROCESS PDF
# =========================
def process_pdf(file_bytes, progress_box):

    pages = convert_from_bytes(file_bytes, dpi=150)

    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"

    ws.append(["STT", "SM", "Ngày", "Trang"])

    total = len(pages)

    found = 0

    for i, img in enumerate(pages, start=1):

        percent = int(i / total * 100)
        progress_box.markdown(f"⏳ {percent}% ({i}/{total})")

        sm, date = ocr(img)

        if sm and date:
            ws.append([found + 1, sm, date, i])
            found += 1

    # =========================
    # FIX FORMAT EXCEL
    # =========================
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for c in row:
            c.border = border
            if row[0].row == 1:
                c.font = Font(bold=True)

    return wb, found


# =========================
# UI
# =========================
uploaded = st.file_uploader("📂 Chọn file PDF", type=["pdf"])

if uploaded and st.button("🚀 Bắt đầu xử lý"):

    progress_box = st.empty()

    file_bytes = uploaded.read()

    with st.spinner("Đang OCR..."):

        wb, count = process_pdf(file_bytes, progress_box)

        # =========================
        # SAVE TO MEMORY (KHÔNG LỖI DOWNLOAD)
        # =========================
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

    # =========================
    # RESULT
    # =========================
    if count == 0:
        st.error("❌ KHÔNG ĐỌC ĐƯỢC DATA TRONG PDF")
        st.warning("👉 PDF có thể là scan mờ hoặc format khác SM / DATE")
    else:
        st.success(f"🎉 DONE - LẤY ĐƯỢC {count} DÒNG DATA")

        st.download_button(
            "📥 TẢI EXCEL",
            data=buffer,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
