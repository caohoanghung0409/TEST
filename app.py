import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
import pandas as pd
import re
import time
import numpy as np
import cv2
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF → EXCEL FAST", layout="wide")
st.title("🚀 THL PDF → EXCEL (FAST VERSION)")

# =========================
# REGEX
# =========================
SM_REGEX = re.compile(r"SM\s*[-:]?\s*\d{4}\s*\.?\s*\d{4}")
DATE_REGEX = re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}")

# =========================
# PREPROCESS (NHẸ + NHANH)
# =========================
def preprocess(img):

    img = np.array(img)

    img = cv2.resize(img, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    gray = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31, 2
    )

    return gray

# =========================
# OCR WORKER (PARALLEL)
# =========================
def ocr_worker(args):

    img, page = args

    img = preprocess(img)

    text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return (page, sm.group(), date.group())

    return None

# =========================
# PROCESS PDF
# =========================
def process_pdf(file_bytes, progress_box):

    # 🔥 giảm DPI để tăng tốc
    pages = convert_from_bytes(file_bytes, dpi=100)

    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"

    ws.append(["STT", "SM", "Ngày", "Trang"])

    total = len(pages)
    start = time.time()

    results = []

    # =========================
    # PARALLEL OCR (QUAN TRỌNG)
    # =========================
    with ThreadPoolExecutor(max_workers=6) as executor:

        tasks = [(img, i) for i, img in enumerate(pages, start=1)]

        for idx, result in enumerate(executor.map(ocr_worker, tasks), start=1):

            percent = int(idx / total * 100)

            elapsed = time.time() - start
            speed = idx / elapsed if elapsed else 0
            eta = int((total - idx) / speed) if speed else 0

            progress_box.markdown(f"⚡ {percent}% | ETA {eta}s")

            if result:
                results.append(result)

    # =========================
    # WRITE EXCEL
    # =========================
    for i, r in enumerate(results, start=1):
        ws.append([i, r[1], r[2], r[0]])

    # nếu không có data
    if not results:
        ws.append(["KHÔNG CÓ DATA", "", "", ""])

    # =========================
    # FORMAT EXCEL
    # =========================
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for c in row:
            c.border = border
            if row[0].row == 1:
                c.font = Font(bold=True)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer, len(results)

# =========================
# UI
# =========================
file = st.file_uploader("📂 Chọn PDF", type=["pdf"])

if file and st.button("🚀 BẮT ĐẦU"):

    progress_box = st.empty()

    file_bytes = file.read()

    with st.spinner("Đang xử lý OCR nhanh..."):

        result, count = process_pdf(file_bytes, progress_box)

    # =========================
    # RESULT
    # =========================
    if count == 0:
        st.error("❌ KHÔNG ĐỌC ĐƯỢC DATA (PDF có thể scan mờ)")
    else:
        st.success(f"🎉 DONE - {count} dòng data")

        st.download_button(
            "📥 TẢI EXCEL",
            data=result,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
