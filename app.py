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

st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# OCR
# =========================
def ocr_extract(img):

    def normalize(text):
        text = text.upper()
        text = text.replace("P8", "PR")
        text = text.replace("S0", "SO")
        text = text.replace(" ", "")
        return text

    def is_valid_header(text):
        t = text.upper()

        return (
            "NHUA THIEU NIEN TIEN PHONG" in t
            and "NHUATIENPHONG.VN" in t
        )

    def extract_from_text(text):

        if not is_valid_header(text):
            return None, None, None

        lines = text.split("\n")

        sm = None
        prso = None
        date = None

        for line in lines:
            raw = line.strip()
            clean = normalize(raw)

            # SM
            if not sm:
                m = re.search(r"(SM\d{4}\.\d{4})", clean)
                if m:
                    sm = m.group(1)

            # PR/SO
            if not prso:
                m = re.search(r"(PR\d{4}\.\d{4}/SO\d{4}\.\d{4})", clean)
                if m:
                    prso = m.group(1)

            # PR riêng
            if not prso:
                m = re.search(r"(PR\d{4}\.\d{4})", clean)
                if m:
                    prso = m.group(1)

            # SO fallback
            if not prso:
                m = re.search(r"(SO\d{4}\.\d{4})", clean)
                if m:
                    prso = m.group(1)

            # DATE
            if not date:
                d = re.search(r"(\d{2}/\d{2}/\d{4})", raw)
                if d:
                    date = d.group(1)

        return sm, prso, date

    w, h = img.size

    # ===== 1. CHECK HEADER NHANH (rất quan trọng) =====
    header = img.crop((0, 0, w, int(h * 0.25)))

    text_quick = pytesseract.image_to_string(
        header, lang='eng', config='--oem 3 --psm 6'
    )

    if not is_valid_header(text_quick):
        return None, None, None  # ❗ skip luôn

    # ===== 2. OCR CHÍNH (ít lần) =====
    for variant in [
        img.crop((0, 0, w, int(h * 0.4))),
        img
    ]:
        text = pytesseract.image_to_string(
            variant, lang='eng', config='--oem 3 --psm 6'
        )

        sm, prso, date = extract_from_text(text)

        if (sm or prso) and date:
            return sm, prso, date

    # fallback rotate 180
    img2 = img.rotate(180, expand=True)
    text = pytesseract.image_to_string(
        img2, lang='eng', config='--oem 3 --psm 6'
    )

    return extract_from_text(text)


# =========================
# PROCESS
# =========================
def extract_pdf(file):

    results = []
    images = convert_from_bytes(file.read(), dpi=150)

    for i, img in enumerate(images, start=1):

        sm, prso, date = ocr_extract(img)

        if sm or prso:
            results.append({
                "SM": sm if sm else "",
                "PR/SO": prso if prso else "",
                "Ngày": date if date else "",
                "Trang": i
            })

    return results


# =========================
# UI
# =========================
st.title("🚀 THL PDF → EXCEL")

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý"):

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for f in uploaded_files:

                data = extract_pdf(f)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # format excel
        wb = load_workbook(tmp_excel.name)

        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border

            for cell in ws[1]:
                cell.font = Font(bold=True)

        wb.save(tmp_excel.name)

        st.success("🎉 XONG!")

        with open(tmp_excel.name, "rb") as f:
            st.download_button(
                "📥 Tải Excel",
                f,
                file_name="ket_qua.xlsx"
            )
