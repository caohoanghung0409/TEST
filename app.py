import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import base64
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

st.markdown("## 🚀 THL PDF → EXCEL (FIX MẤT SM / PR / SO)")

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# CLEAN FUNCTION
# =========================
def clean_code(x):
    if not x:
        return None
    return re.sub(r"\s+", "", x)

# =========================
# OCR FUNCTION (ROBUST)
# =========================
def process_page(img):

    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )

    # 🔥 FIX regex chịu lỗi OCR
    sm = re.search(r"(SM\s*\d{3,5}[\.\s]?\d{3,5})", text)
    pr = re.search(r"(PR\s*\d{3,5}[\.\s]?\d{3,5})", text)
    so = re.search(r"(SO\s*\d{3,5}[\.\s]?\d{3,5})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    return (
        clean_code(sm.group(1)) if sm else None,
        clean_code(pr.group(1)) if pr else None,
        clean_code(so.group(1)) if so else None,
        date.group(1) if date else None
    )

# =========================
# EXTRACT PDF (FIX MẤT DỮ LIỆU)
# =========================
def extract_pdf(file):

    results = []

    # 🔥 FIX QUAN TRỌNG: không dùng file.read()
    pdf_bytes = file.getvalue()

    # giảm DPI để tránh treo + nhanh hơn
    images = convert_from_bytes(pdf_bytes, dpi=120)

    # 🔥 GIỮ STATE XUYÊN TRANG
    last_sm = None
    last_pr = None
    last_so = None

    for img in images:

        # ❗ KHÔNG CROP CỨNG (tránh mất dữ liệu)
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.8)))  # chỉ cắt nhẹ

        sm, pr, so, date = process_page(img)

        # 🔥 GIỮ GIÁ TRỊ QUA TRANG
        if sm:
            last_sm = sm
        if pr:
            last_pr = pr
        if so:
            last_so = so

        # chỉ cần có data là ghi
        if last_sm or last_pr or last_so:

            results.append({
                "SM": last_sm,
                "PR": last_pr,
                "SO": last_so,
                "Ngày": date
            })

    return results

# =========================
# RUN BUTTON
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý", type="primary"):

        with st.spinner("⏳ Đang OCR PDF... vui lòng chờ"):

            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

            with pd.ExcelWriter(excel_file.name, engine='openpyxl') as writer:

                for f in uploaded_files:

                    data = extract_pdf(f)

                    df = pd.DataFrame(data)

                    if not df.empty:
                        df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # =========================
            # AUTO WIDTH EXCEL
            # =========================
            wb = load_workbook(excel_file.name)

            for ws in wb.worksheets:
                for col in ws.columns:
                    max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                    ws.column_dimensions[col[0].column_letter].width = max_len + 3

            wb.save(excel_file.name)

            st.session_state.excel = excel_file.name
            st.success("🎉 Xử lý xong!")

# =========================
# DOWNLOAD
# =========================
if "excel" in st.session_state and st.session_state.excel:

    with open(st.session_state.excel, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
        <a href="data:application/octet-stream;base64,{b64}" download="result.xlsx">
            📥 Tải file Excel
        </a>
    """, unsafe_allow_html=True)
