import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import base64
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")
st.title("🚀 THL PDF → EXCEL (STABLE VERSION)")

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Upload PDF",
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
# OCR + ANCHOR
# =========================
def extract_from_page(img):

    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # =========================
    # FIND ANCHOR
    # =========================
    start_idx = -1
    for i, line in enumerate(lines):
        if "PHIẾU GIAO HÀNG" in line.upper():
            start_idx = i
            break

    if start_idx != -1:
        lines = lines[start_idx:]

    zone_text = " ".join(lines)

    # =========================
    # EXTRACT
    # =========================
    sm = re.search(r"(SM\s*\d{3,5}[\.\s]?\d{3,5})", zone_text)
    pr = re.search(r"(PR\s*\d{3,5}[\.\s]?\d{3,5})", zone_text)
    so = re.search(r"(SO\s*\d{3,5}[\.\s]?\d{3,5})", zone_text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", zone_text)

    return {
        "SM": clean(sm.group(1)) if sm else None,
        "PR": clean(pr.group(1)) if pr else None,
        "SO": clean(so.group(1)) if so else None,
        "Ngày": date.group(1) if date else None
    }

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file):

    pdf_bytes = file.getvalue()
    images = convert_from_bytes(pdf_bytes, dpi=120)

    results = []

    for img in images:

        row = extract_from_page(img)

        # chỉ append nếu có ít nhất 1 data
        if any(row.values()):
            results.append(row)

    return results

# =========================
# RUN BUTTON
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý", type="primary"):

        all_data = []
        has_data = False

        for f in uploaded_files:

            data = extract_pdf(f)

            if data:
                has_data = True
                df = pd.DataFrame(data)

                if not df.empty:
                    df.insert(0, "STT", range(1, len(df)+1))
                    df["FILE"] = f.name

                all_data.append((f.name, df))

        # =========================
        # ANTI EMPTY FILE CRASH
        # =========================
        if not has_data:
            st.error("❌ Không tìm thấy dữ liệu trong PDF")
            st.stop()

        # =========================
        # CREATE EXCEL SAFE
        # =========================
        excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(excel_file.name, engine='openpyxl') as writer:

            wrote = False

            for name, df in all_data:

                if df is not None and not df.empty:
                    wrote = True
                    df.to_excel(writer, sheet_name=name[:31], index=False)

            # 🔥 GUARANTEE SHEET (ANTI CRASH OPENPYXL)
            if not wrote:
                pd.DataFrame([{"ERROR": "NO VALID DATA"}]).to_excel(writer, index=False)

        # =========================
        # AUTO WIDTH
        # =========================
        wb = load_workbook(excel_file.name)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

        wb.save(excel_file.name)

        # =========================
        # DOWNLOAD
        # =========================
        with open(excel_file.name, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()

        st.success("🎉 Xử lý xong!")

        st.markdown(f"""
        <a href="data:application/octet-stream;base64,{b64}" download="result.xlsx">
            📥 Download Excel
        </a>
        """, unsafe_allow_html=True)
