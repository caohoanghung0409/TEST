import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import base64
from openpyxl import load_workbook

st.set_page_config(page_title="PDF FIX TABLE MODE", layout="wide")
st.title("🚀 FIX PR / SO / SM LẶP & THIẾU")

uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

# =========================
# CLEAN
# =========================
def clean(x):
    if not x:
        return None
    return re.sub(r"\s+", "", x)

# =========================
# PARSE LINE BY LINE (QUAN TRỌNG)
# =========================
def extract_pdf(file):

    pdf_bytes = file.getvalue()
    images = convert_from_bytes(pdf_bytes, dpi=120)

    results = []

    # 🔥 STATE MACHINE
    current_sm = None
    current_pr = None
    current_so = None

    for img in images:

        text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")
        lines = [l.strip() for l in text.split("\n") if l.strip()]

        for line in lines:

            # =========================
            # UPDATE HEADER VALUES
            # =========================
            sm = re.search(r"SM\s*\d{3,5}[\.\s]?\d{3,5}", line)
            pr = re.search(r"PR\s*\d{3,5}[\.\s]?\d{3,5}", line)
            so = re.search(r"SO\s*\d{3,5}[\.\s]?\d{3,5}", line)
            date = re.search(r"\d{2}/\d{2}/\d{4}", line)

            if sm:
                current_sm = clean(sm.group())
            if pr:
                current_pr = clean(pr.group())
            if so:
                current_so = clean(so.group())

            # =========================
            # ITEM ROW (DATA LINE)
            # =========================
            if any([current_pr, current_so, current_sm]):

                results.append({
                    "SM": current_sm,
                    "PR": current_pr,
                    "SO": current_so,
                    "Ngày": date.group() if date else None,
                    "Raw": line
                })

    return results

# =========================
# RUN
# =========================
if uploaded_files:

    if st.button("🚀 Xử lý"):

        all_data = []

        for f in uploaded_files:

            data = extract_pdf(f)

            df = pd.DataFrame(data)

            if not df.empty:
                df.insert(0, "STT", range(1, len(df)+1))

            all_data.append((f.name, df))

        # =========================
        # SAFE EXPORT
        # =========================
        excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(excel_file.name, engine='openpyxl') as writer:

            wrote = False

            for name, df in all_data:

                if df is not None and not df.empty:
                    wrote = True
                    df.to_excel(writer, sheet_name=name[:31], index=False)

            if not wrote:
                pd.DataFrame([{"ERROR": "NO DATA FOUND"}]).to_excel(writer, index=False)

        # =========================
        # DOWNLOAD
        # =========================
        with open(excel_file.name, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()

        st.success("DONE")

        st.markdown(f"""
        <a href="data:application/octet-stream;base64,{b64}" download="result.xlsx">
        📥 Download Excel
        </a>
        """, unsafe_allow_html=True)
