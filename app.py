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
# OCR xoay 4 hướng
# =========================
def process_page_multi_angle(img):
    angles = [0, 90, 180, 270]

    for angle in angles:
        rotated = img.rotate(angle, expand=True)

        # crop 40% phía trên
        w, h = rotated.size
        cropped = rotated.crop((0, 0, w, int(h * 0.4)))

        text = pytesseract.image_to_string(
            cropped,
            lang='eng',
            config='--oem 3 --psm 6'
        )

        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

        if sm and date:
            return sm.group(1), date.group(1)

    return None, None

# =========================
# GLOBAL BAR
# =========================
def render_global_bar(percent):
    return f"""
<div style="margin:10px 0;">
    <div style="background:#ddd;border-radius:10px;overflow:hidden;">
        <div style="width:{percent}%;background:#4CAF50;color:white;text-align:center;">
            {percent}%
        </div>
    </div>
</div>
"""

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages_all):

    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        processed_pages[0] += 1

        global_percent = int((processed_pages[0] / total_pages_all) * 100)
        global_box.markdown(render_global_bar(global_percent), unsafe_allow_html=True)

        box.write(f"📄 {file.name} - Trang {i}/{total_pages}")

        # FIX CHÍNH Ở ĐÂY
        sm, date = process_page_multi_angle(img)

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

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start_time = time.time()

        total_pages_all = sum(len(convert_from_bytes(f.read(), dpi=50)) for f in uploaded_files)
        for f in uploaded_files:
            f.seek(0)

        processed_pages = [0]

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for i, f in enumerate(uploaded_files):

                data = extract_pdf(
                    f, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

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
