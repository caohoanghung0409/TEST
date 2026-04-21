import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import tempfile
import os
import time
import base64
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# SESSION
if "processing" not in st.session_state:
    st.session_state.processing = False
if "done" not in st.session_state:
    st.session_state.done = False
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

st.markdown("## 🚀 THL PDF → EXCEL (SIÊU NHANH)")

uploaded_files = st.file_uploader("📂 Chọn file PDF", type=["pdf"], accept_multiple_files=True)

# =========================
# TEXT EXTRACT (KHÔNG OCR)
# =========================
def extract_pdf_text(file):

    results = []
    doc = fitz.open(stream=file.read(), filetype="pdf")

    for i, page in enumerate(doc, start=1):

        text = page.get_text()

        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

        if sm and date:
            results.append({
                "SM": sm.group(1),
                "Ngày": date.group(1),
                "Trang": i
            })

    return results

# =========================
# PROCESS
# =========================
if uploaded_files:

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start = time.time()

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        has_data = False

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for f in uploaded_files:

                data = extract_pdf_text(f)

                if data:
                    has_data = True
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            if not has_data:
                df = pd.DataFrame([{"Thông báo": "Không có dữ liệu hợp lệ"}])
                df.to_excel(writer, sheet_name="NO_DATA", index=False)

        # format excel
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

        st.success(f"🎉 Xong trong {round(time.time()-start,2)}s")
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    with open(st.session_state.excel_file, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
    <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 Xử lý file mới"):
        st.session_state.done = False
        st.rerun()
