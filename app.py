import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import time
import base64
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font
from concurrent.futures import ProcessPoolExecutor
import cv2
import numpy as np

st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

st.title("🚀 THL PDF → EXCEL (MAX SPEED)")

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# OCR SIÊU NHANH
# =========================
def ocr_fast(img):

    # 👉 convert PIL -> OpenCV
    img = np.array(img)

    h, w = img.shape[:2]

    # 👉 crop vùng trên
    img = img[0:int(h*0.35), :]

    # 👉 resize nhỏ lại (giảm load)
    img = cv2.resize(img, None, fx=0.7, fy=0.7)

    # 👉 grayscale
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 👉 threshold tăng độ rõ chữ
    _, img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY)

    def read(im):
        text = pytesseract.image_to_string(
            im,
            lang='eng',
            config='--oem 3 --psm 6'
        )
        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
        return sm, date

    # 👉 chỉ 2 hướng
    for v in [img, cv2.rotate(img, cv2.ROTATE_180)]:
        sm, date = read(v)
        if sm and date:
            return sm.group(1), date.group(1)

    return None, None


# =========================
# XỬ LÝ 1 TRANG (song song)
# =========================
def process_page(args):
    img, index = args

    sm, date = ocr_fast(img)

    if sm and date:
        return {
            "SM": sm,
            "Ngày": date,
            "Trang": index + 1
        }
    return None


# =========================
# MAIN
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý"):

        start = time.time()

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for f in uploaded_files:

                st.write(f"📄 Đang xử lý: {f.name}")

                # 🔥 DPI thấp + nhanh
                images = convert_from_bytes(f.read(), dpi=85)

                # 👉 chạy song song
                with ProcessPoolExecutor() as executor:
                    results = list(executor.map(
                        process_page,
                        [(img, i) for i, img in enumerate(images)]
                    ))

                # lọc None
                results = [r for r in results if r]

                if results:
                    df = pd.DataFrame(results)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # =========================
        # FORMAT EXCEL
        # =========================
        wb = load_workbook(tmp_excel.name)

        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:

            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 2

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border

            for cell in ws[1]:
                cell.font = Font(bold=True)

        wb.save(tmp_excel.name)

        # =========================
        # DOWNLOAD
        # =========================
        with open(tmp_excel.name, "rb") as f:
            data = f.read()

        b64 = base64.b64encode(data).decode()

        st.success(f"🎉 Xong trong {round(time.time()-start,2)}s")

        st.markdown(f"""
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="result.xlsx">
        📥 Tải file Excel
        </a>
        """, unsafe_allow_html=True)
