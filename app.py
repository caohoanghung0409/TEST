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
st.markdown('<div style="font-size:22px;font-weight:700;">🚀 THL PDF → EXCEL</div>', unsafe_allow_html=True)

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
# OCR (FIX XOAY THÔNG MINH)
# =========================
def process_page(img):
    # thử trước không xoay (nhanh)
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    if sm and date:
        return sm.group(1), date.group(1)

    # nếu fail → mới xoay
    for angle in [90, 180, 270]:
        rotated = img.rotate(angle, expand=True)

        text = pytesseract.image_to_string(rotated, lang='eng', config='--oem 3 --psm 6')

        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

        if sm and date:
            return sm.group(1), date.group(1)

    return None, None

# =========================
# GLOBAL BAR
# =========================
def render_global_bar(percent, speed, eta):
    return f"""
<div style="margin:15px 0;">
    <div style="display:flex;justify-content:space-between;font-size:13px;">
        <div>⚡ {percent}%</div>
        <div>🚀 {speed:.2f} pages/s • ⏳ {eta}s</div>
    </div>
    <div style="height:20px;background:#e5e7eb;border-radius:999px;overflow:hidden;">
        <div style="width:{percent}%;height:100%;background:linear-gradient(90deg,#3b82f6,#22c55e);"></div>
    </div>
</div>
"""

# =========================
# EXTRACT PDF
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages_all):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        processed_pages[0] += 1

        percent = int((i/total_pages)*100)
        global_percent = int((processed_pages[0] / total_pages_all) * 100)

        elapsed = time.time() - start_time
        speed = processed_pages[0] / elapsed if elapsed > 0 else 0
        remaining = total_pages_all - processed_pages[0]
        eta = int(remaining / speed) if speed > 0 else 0

        global_box.markdown(render_global_bar(global_percent, speed, eta), unsafe_allow_html=True)

        box.markdown(f"""
📄 {file.name} — Trang {i}/{total_pages} ({percent}%)
""")

        # crop
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({
                "Trang": i,   # ✅ thêm số trang
                "SM": sm,
                "Ngày": date
            })

    return results

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

        st.write("⏳ Đang xử lý...")

        start_time = time.time()

        total_pages_all = sum(len(convert_from_bytes(f.read(), dpi=50)) for f in uploaded_files)
        for f in uploaded_files:
            f.seek(0)

        processed_pages = [0]

        excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(excel_tmp.name, engine="openpyxl") as writer:

            for i, f in enumerate(uploaded_files):

                data = extract_pdf(
                    f, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # auto width
        wb = load_workbook(excel_tmp.name)
        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3
        wb.save(excel_tmp.name)

        st.session_state.excel_file = excel_tmp.name
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
        <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
        style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
