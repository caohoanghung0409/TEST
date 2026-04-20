import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
import pandas as pd
import re
import tempfile
import time
import base64
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font

st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# REGEX
# =========================
SM_REGEX = re.compile(r"(SM\d{4}\.\d{4})")
DATE_REGEX = re.compile(r"(\d{2}/\d{2}/\d{4})")

# =========================
# OCR
# =========================
def ocr_extract(img):
    text = pytesseract.image_to_string(img, lang="eng", config="--oem 3 --psm 6")

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return sm.group(1), date.group(1)

    return None, None

# =========================
# PDF PROCESS
# =========================
def process_pdf(file_bytes, file_name, ws, global_counter, total_pages_all, global_box, start_time):

    pages = convert_from_bytes(file_bytes, dpi=110)
    results = []

    for i, img in enumerate(pages, start=1):

        global_counter[0] += 1

        percent = int(global_counter[0] / total_pages_all * 100)
        elapsed = time.time() - start_time
        speed = global_counter[0] / elapsed if elapsed else 0
        eta = int((total_pages_all - global_counter[0]) / speed) if speed else 0

        global_box.markdown(f"⚡ {percent}% | ETA {eta}s")

        sm, date = ocr_extract(img)

        if sm and date:
            results.append((sm, date, i))

    # =========================
    # WRITE SHEET (OPENPYXL DIRECT)
    # =========================
    if not results:
        ws.append(["THÔNG BÁO", "KHÔNG CÓ DỮ LIỆU", ""])
        return

    ws.append(["STT", "SM", "Ngày", "Trang"])

    for idx, row in enumerate(results, 1):
        ws.append([idx, row[0], row[1], row[2]])

# =========================
# UI
# =========================
st.title("🚀 THL PDF → EXCEL (FIXED VERSION)")

files = st.file_uploader("Chọn PDF", type=["pdf"], accept_multiple_files=True)

if files and st.button("Bắt đầu xử lý"):

    start_time = time.time()
    global_counter = [0]

    wb = Workbook()
    wb.remove(wb.active)  # xóa sheet mặc định

    global_box = st.empty()

    total_pages_all = 0
    file_bytes_list = []

    # tính tổng page (NHANH)
    for f in files:
        b = f.getvalue()
        file_bytes_list.append((f.name, b))
        total_pages_all += int(pdfinfo_from_bytes(b)["Pages"])

    # xử lý từng file
    for name, b in file_bytes_list:

        ws = wb.create_sheet(title=name[:31])

        process_pdf(
            b,
            name,
            ws,
            global_counter,
            total_pages_all,
            global_box,
            start_time
        )

    # format Excel
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                if row[0].row == 1:
                    cell.font = Font(bold=True)

    # save file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)

    with open(tmp.name, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    st.success("Hoàn thành!")

    st.download_button(
        "📥 Download Excel",
        data=f,
        file_name="output.xlsx"
    )
