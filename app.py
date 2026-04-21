import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import base64
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# SESSION
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False
if "done" not in st.session_state:
    st.session_state.done = False
if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

# =========================
# UI
# =========================
st.title("🚀 THL PDF → EXCEL")

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# OCR CORE
# =========================
def ocr_extract(img):

    def normalize(text):
        text = text.upper()
        text = text.replace("P8", "PR")
        text = text.replace("S0", "SO")
        return text

    def is_valid_header(text):
        t = text.upper()
        return "TIEN PHONG" in t or "NHUATIENPHONG" in t

    def extract_from_text(text):

        if not is_valid_header(text):
            return None, None, None

        sm = None
        prso = None
        date = None

        lines = text.split("\n")

        for line in lines:
            raw = line.strip()
            clean = normalize(raw)

            # ===== SM =====
            if not sm:
                m = re.search(r"(SM\d{4}\.\d{4})", clean)
                if m:
                    sm = m.group(1)

            # ===== PR + SO GHÉP =====
            if not prso:

                pr_match = re.search(r"P[R]\d{4}\.\d{4}", clean)
                so_match = re.search(r"S[O]\d{4}\.\d{4}", clean)

                if pr_match and so_match:
                    pr_val = pr_match.group(0)
                    so_val = so_match.group(0)
                    prso = f"{pr_val}/{so_val}"

            # ===== PR riêng =====
            if not prso:
                m = re.search(r"P[R]\d{4}\.\d{4}", clean)
                if m:
                    prso = m.group(0)

            # ===== SO fallback =====
            if not prso:
                m = re.search(r"S[O]\d{4}\.\d{4}", clean)
                if m:
                    prso = m.group(0)

            # ===== DATE =====
            if not date:
                d = re.search(r"(\d{2}/\d{2}/\d{4})", raw)
                if d:
                    date = d.group(1)

        return sm, prso, date

    w, h = img.size

    # ===== CHECK HEADER =====
    header = img.crop((0, 0, w, int(h * 0.25)))
    quick_text = pytesseract.image_to_string(header, config='--oem 3 --psm 6')

    if not is_valid_header(quick_text):
        return None, None, None

    # ===== OCR =====
    for variant in [
        img.crop((0, 0, w, int(h * 0.4))),
        img
    ]:
        text = pytesseract.image_to_string(variant, config='--oem 3 --psm 6')
        sm, prso, date = extract_from_text(text)

        if (sm or prso) and date:
            return sm, prso, date

    return None, None, None

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
# MAIN
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý"):

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            has_data = False

            for f in uploaded_files:

                data = extract_pdf(f)

                if data:
                    has_data = True

                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            if not has_data:
                df = pd.DataFrame([{"Thông báo": "Không tìm thấy dữ liệu"}])
                df.to_excel(writer, sheet_name="KET_QUA", index=False)

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

        st.success("🎉 HOÀN THÀNH")

        with open(tmp_excel.name, "rb") as f:
            st.download_button("📥 Tải Excel", f, file_name="ket_qua.xlsx")
