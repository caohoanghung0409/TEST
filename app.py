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

# =========================
# SESSION STATE
# =========================
if "done" not in st.session_state:
    st.session_state.done = False

if "excel" not in st.session_state:
    st.session_state.excel = None

# =========================
# UI STYLE (GIỮ ĐƠN GIẢN + ỔN ĐỊNH)
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem;}
.stApp { background: #f1f5f9; }

div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:12px 20px;
    font-weight:600;
}
</style>
""", unsafe_allow_html=True)

st.markdown("## 🚀 THL PDF → EXCEL (SM / PR / SO)")

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# OCR FUNCTION
# =========================
def process_page(img):
    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )

    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    pr = re.search(r"(PR\d{4}\.\d{4})", text)
    so = re.search(r"(SO\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    return (
        sm.group(1) if sm else None,
        pr.group(1) if pr else None,
        so.group(1) if so else None,
        date.group(1) if date else None
    )

# =========================
# EXTRACT PDF (FIX CHÍNH Ở ĐÂY)
# =========================
def extract_pdf(file):

    results = []

    # 🔥 FIX QUAN TRỌNG: KHÔNG DÙNG file.read()
    pdf_bytes = file.getvalue()

    # giảm dpi để tránh treo
    images = convert_from_bytes(pdf_bytes, dpi=120)

    for img in images:

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, pr, so, date = process_page(img)

        if sm or pr or so:
            results.append({
                "SM": sm,
                "PR": pr,
                "SO": so,
                "Ngày": date
            })

    return results

# =========================
# RUN BUTTON
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý", type="primary"):

        with st.spinner("⏳ Đang xử lý PDF... vui lòng chờ"):

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
            st.session_state.done = True
            st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 Xử lý hoàn tất!")

    with open(st.session_state.excel, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
        <a href="data:application/octet-stream;base64,{b64}" download="result.xlsx">
            📥 Tải file Excel
        </a>
    """, unsafe_allow_html=True)

    if st.button("🔄 Xử lý file mới"):
        st.session_state.done = False
        st.rerun()
