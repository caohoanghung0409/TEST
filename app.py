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

st.markdown("## 🚀 THL PDF → EXCEL (FIX DUPLICATE + SM PR SO)")

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# CLEAN
# =========================
def clean(x):
    if not x:
        return None
    return re.sub(r"\s+", "", x)

# =========================
# OCR
# =========================
def extract_text(img):

    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )

    sm = re.search(r"(SM\s*\d{3,5}[\.\s]?\d{3,5})", text)
    pr = re.search(r"(PR\s*\d{3,5}[\.\s]?\d{3,5})", text)
    so = re.search(r"(SO\s*\d{3,5}[\.\s]?\d{3,5})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    return (
        clean(sm.group(1)) if sm else None,
        clean(pr.group(1)) if pr else None,
        clean(so.group(1)) if so else None,
        date.group(1) if date else None
    )

# =========================
# CORE FIX (QUAN TRỌNG)
# =========================
def extract_pdf(file):

    results = []

    # 🔥 FIX: dùng getvalue()
    pdf_bytes = file.getvalue()

    # convert PDF → images
    images = convert_from_bytes(pdf_bytes, dpi=120)

    for img in images:

        # chỉ crop nhẹ (không cắt mất data)
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.8)))

        sm, pr, so, date = extract_text(img)

        # 🔥 QUAN TRỌNG: KHÔNG carry forward
        # mỗi page = 1 record độc lập

        if sm or pr or so or date:

            results.append({
                "SM": sm,
                "PR": pr,
                "SO": so,
                "Ngày": date
            })

    return results

# =========================
# RUN
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý", type="primary"):

        with st.spinner("⏳ Đang OCR PDF..."):

            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

            with pd.ExcelWriter(excel_file.name, engine='openpyxl') as writer:

                for f in uploaded_files:

                    data = extract_pdf(f)

                    df = pd.DataFrame(data)

                    # thêm STT đúng từng sheet
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

            st.session_state.file = excel_file.name
            st.success("🎉 Xử lý xong!")

# =========================
# DOWNLOAD
# =========================
if "file" in st.session_state:

    with open(st.session_state.file, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
        <a href="data:application/octet-stream;base64,{b64}" download="result.xlsx">
            📥 Tải Excel
        </a>
    """, unsafe_allow_html=True)
