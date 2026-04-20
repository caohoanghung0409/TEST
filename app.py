import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import base64
from openpyxl import load_workbook

st.set_page_config(page_title="PDF TO EXCEL FIX DUP", layout="wide")

st.title("🚀 FIX TRÙNG PR / SO / SM")

uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

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
def parse_text(text):

    sm = re.search(r"SM\s*\d{3,5}[\.\s]?\d{3,5}", text)
    pr = re.search(r"PR\s*\d{3,5}[\.\s]?\d{3,5}", text)
    so = re.search(r"SO\s*\d{3,5}[\.\s]?\d{3,5}", text)
    date = re.search(r"\d{2}/\d{2}/\d{4}", text)

    return (
        clean(sm.group()) if sm else None,
        clean(pr.group()) if pr else None,
        clean(so.group()) if so else None,
        date.group() if date else None
    )

# =========================
# MAIN FIX ENGINE
# =========================
def extract_pdf(file):

    pdf_bytes = file.getvalue()
    images = convert_from_bytes(pdf_bytes, dpi=120)

    results = []

    for img in images:

        text = pytesseract.image_to_string(img, config='--oem 3 --psm 6')

        sm, pr, so, date = parse_text(text)

        # 🔥 CHỐNG TRÙNG TRONG CÙNG FILE
        key = (pr, so, date)

        # nếu đã tồn tại → skip
        if any(r.get("key") == key for r in results):
            continue

        results.append({
            "SM": sm,
            "PR": pr,
            "SO": so,
            "Ngày": date,
            "key": key
        })

    # bỏ key trước khi export
    for r in results:
        r.pop("key", None)

    return results

# =========================
# RUN
# =========================
if uploaded_files:

    if st.button("🚀 Xử lý"):

        excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(excel_file.name, engine='openpyxl') as writer:

            for f in uploaded_files:

                data = extract_pdf(f)

                df = pd.DataFrame(data)

                if not df.empty:
                    df.insert(0, "STT", range(1, len(df)+1))

                df.to_excel(writer, sheet_name=f.name[:31], index=False)

        wb = load_workbook(excel_file.name)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

        wb.save(excel_file.name)

        with open(excel_file.name, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()

        st.success("DONE")

        st.markdown(f"""
        <a href="data:application/octet-stream;base64,{b64}" download="result.xlsx">
        📥 Download Excel
        </a>
        """, unsafe_allow_html=True)
