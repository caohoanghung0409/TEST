import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
import re
import numpy as np
import cv2
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font
import time

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF → EXCEL FAST", layout="wide")
st.title("🚀 THL PDF → EXCEL (FAST + STABLE)")

# =========================
# REGEX (LINH HOẠT)
# =========================
SM_REGEX = re.compile(r"SM\s*[-:]?\s*\d{4}\s*\.?\s*\d{4}")
DATE_REGEX = re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}")

# =========================
# PREPROCESS NHẸ (QUAN TRỌNG SPEED)
# =========================
def preprocess(img):
    img = np.array(img)

    # nhẹ hơn resize lớn để tăng tốc
    img = cv2.resize(img, None, fx=1.2, fy=1.2)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # threshold đơn giản (nhanh hơn adaptive)
    gray = cv2.threshold(gray, 160, 255, cv2.THRESH_BINARY)[1]

    return gray

# =========================
# OCR
# =========================
def ocr(img):
    img = preprocess(img)

    text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return sm.group(), date.group()

    return None, None

# =========================
# PROCESS (STREAMING - KHÔNG LOAD FULL RAM)
# =========================
def process(file_bytes, progress_box):

    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"

    ws.append(["STT", "SM", "Ngày", "Trang"])

    info = pdfinfo_from_bytes(file_bytes)
    total_pages = int(info["Pages"])

    start = time.time()
    found = 0

    # 🔥 xử lý từng page (tránh load all images → nhanh + ổn định)
    for i in range(1, total_pages + 1):

        page = convert_from_bytes(
            file_bytes,
            dpi=100,
            first_page=i,
            last_page=i
        )[0]

        sm, date = ocr(page)

        if sm and date:
            found += 1
            ws.append([found, sm, date, i])

        # progress
        percent = int(i / total_pages * 100)

        elapsed = time.time() - start
        speed = i / elapsed if elapsed else 0
        eta = int((total_pages - i) / speed) if speed else 0

        progress_box.markdown(f"⚡ {percent}% | ETA {eta}s")

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

if file and st.button("🚀 BẮT ĐẦU XỬ LÝ"):

    progress_box = st.empty()

    file_bytes = file.read()

    with st.spinner("Đang xử lý OCR..."):

        result, count = process(file_bytes, progress_box)

    # =========================
    # RESULT
    # =========================
    if count == 0:
        st.error("❌ KHÔNG LẤY ĐƯỢC DATA")
        st.warning("👉 PDF có thể scan mờ hoặc format khác SM / DATE")
    else:
        st.success(f"🎉 HOÀN THÀNH - {count} DÒNG DATA")

        st.download_button(
            "📥 TẢI EXCEL",
            data=result,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
