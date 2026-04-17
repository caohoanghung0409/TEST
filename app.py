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
from PIL import Image

# =========================
# CONFIG
# =========================
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
if "last_uploaded_names" not in st.session_state:
    st.session_state.last_uploaded_names = []
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

# =========================
# HEADER
# =========================
st.markdown("## 🚀 THL PDF → EXCEL")

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

current_names = [f.name for f in uploaded_files] if uploaded_files else []

if current_names != st.session_state.last_uploaded_names:
    st.session_state.processing = False
    st.session_state.done = False
    st.session_state.last_uploaded_names = current_names

# =========================
# OCR SMART (FIX CHÍNH)
# =========================
def ocr_extract(img):

    def read(image):
        text = pytesseract.image_to_string(
            image,
            lang='eng',
            config='--oem 3 --psm 6'
        )
        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
        return sm, date

    w, h = img.size

    # 1. Ảnh gốc FULL
    sm, date = read(img)
    if sm and date:
        return sm.group(1), date.group(1)

    # 2. Ảnh gốc crop
    crop = img.crop((0, 0, w, int(h * 0.4)))
    sm, date = read(crop)
    if sm and date:
        return sm.group(1), date.group(1)

    # 3. Xoay 180 FULL
    img180 = img.rotate(180, expand=True)
    sm, date = read(img180)
    if sm and date:
        return sm.group(1), date.group(1)

    # 4. Xoay 180 crop
    w2, h2 = img180.size
    crop180 = img180.crop((0, 0, w2, int(h2 * 0.4)))
    sm, date = read(crop180)
    if sm and date:
        return sm.group(1), date.group(1)

    # 5. Xoay 90 crop
    img90 = img.rotate(90, expand=True)
    w3, h3 = img90.size
    crop90 = img90.crop((0, 0, w3, int(h3 * 0.4)))
    sm, date = read(crop90)
    if sm and date:
        return sm.group(1), date.group(1)

    # 6. Xoay 270 crop
    img270 = img.rotate(270, expand=True)
    w4, h4 = img270.size
    crop270 = img270.crop((0, 0, w4, int(h4 * 0.4)))
    sm, date = read(crop270)
    if sm and date:
        return sm.group(1), date.group(1)

    return None, None

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file, box):

    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        box.write(f"📄 {file.name} - Trang {i}/{total_pages}")

        sm, date = ocr_extract(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date,
                "Trang": i
            })

    return results

# =========================
# CLEAN SHEET NAME
# =========================
def clean_sheet_name(name):
    name = os.path.splitext(name)[0]
    name = re.sub(r'[\\/*?:\[\]]', '', name)
    return name[:31]

# =========================
# MAIN
# =========================
if uploaded_files:

    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for i, f in enumerate(uploaded_files):

                data = extract_pdf(f, boxes[i])

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = clean_sheet_name(f.name)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

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

        st.session_state.excel_file = tmp_excel.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    with open(st.session_state.excel_file, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
        <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
