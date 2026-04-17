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
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None
if "trigger_download" not in st.session_state:
    st.session_state.trigger_download = False

# =========================
# STYLE (GIỮ NGUYÊN 100%)
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header {
    font-size:22px;
    font-weight:700;
    margin-bottom:10px;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #93c5fd;
    padding: 25px;
    border-radius: 18px;
    background: white;
    transition: 0.3s;
}

div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:12px 24px;
    font-weight:600;
}

.file-row {
    margin-top:12px;
    padding:10px;
    border-radius:12px;
    background:white;
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}

.progress {
    height:8px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
}
.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}

.global-bar {
    height:18px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
}

.global-fill {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 THL PDF → EXCEL</div>', unsafe_allow_html=True)

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.session_state.processing = False
    st.session_state.done = False
    st.session_state.trigger_download = False

# =========================
# OCR
# =========================
def ocr_extract(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1) if sm else None), (date.group(1) if date else None)

# =========================
# PROCESS PDF
# =========================
def process_pdf(file):
    images = convert_from_bytes(file.read(), dpi=150)
    result = []

    for i, img in enumerate(images, start=1):
        sm, date = ocr_extract(img)
        if sm and date:
            result.append({
                "SM": sm,
                "Ngày": date,
                "Trang": i
            })

    return result

# =========================
# RUN
# =========================
if uploaded_files:

    if st.button("🚀 BẮT ĐẦU XỬ LÝ"):

        st.session_state.processing = True

        all_data = []

        for f in uploaded_files:
            data = process_pdf(f)
            if data:
                all_data.append(pd.DataFrame(data))

        final_df = pd.concat(all_data) if all_data else pd.DataFrame()

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False)

        wb = load_workbook(tmp.name)
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

        wb.save(tmp.name)

        st.session_state.excel_file = tmp.name
        st.session_state.processing = False
        st.session_state.done = True
        st.session_state.trigger_download = True
        st.rerun()

# =========================
# AUTO DOWNLOAD (GIỐNG LÚC ĐẦU)
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    with open(st.session_state.excel_file, "rb") as f:
        data = f.read()

    file_name = "THLPDFTOEXCEL.xlsx"
    b64 = base64.b64encode(data).decode()

    # 🔥 AUTO DOWNLOAD TRICK (GIỐNG LÚC ĐẦU BẠN DÙNG)
    if st.session_state.trigger_download:

        st.session_state.trigger_download = False

        st.markdown(f"""
        <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
        style="display:none;"></iframe>
        """, unsafe_allow_html=True)

        st.info("📥 File đang tự động tải về...")

    st.download_button(
        "📥 TẢI LẠI FILE",
        data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.processing = False
        st.rerun()
