import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
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
if "run" not in st.session_state:
    st.session_state.run = False

# =========================
# UI
# =========================
st.markdown("## 🚀 THL PDF → EXCEL")

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# OCR CORE
# =========================
def ocr_extract(img):

    def normalize(text):
        text = text.upper()
        text = text.replace("P8", "PR").replace("S0", "SO")
        text = re.sub(r"\s+", " ", text)
        return text

    def has_anchor(text):
        return "NHUA" in text and "TIEN PHONG" in text

    def extract_line(text):
        pattern = re.compile(
            r"(SM\d{4}\.\d{4}|PR\d{4}\.\d{4}\s*/\s*SO\d{4}\.\d{4}|SO\d{4}\.\d{4}).*?(\d{2}/\d{2}/\d{4})"
        )
        m = pattern.search(text)
        if m:
            code = m.group(1).replace(" ", "")
            date = m.group(2)

            sm, prso = "", ""
            if code.startswith("SM"):
                sm = code
            else:
                prso = code

            return sm, prso, date

        return None, None, None

    w, h = img.size

    for variant in [img, img.rotate(180)]:

        # ===== HEADER =====
        header = variant.crop((0, 0, w, int(h * 0.3)))
        text_header = normalize(pytesseract.image_to_string(header, config='--psm 6'))

        if not has_anchor(text_header):
            continue

        # ===== BODY =====
        body = variant.crop((0, int(h * 0.2), w, int(h * 0.8)))
        raw_body = pytesseract.image_to_string(body, config='--psm 6')
        text_body = normalize(raw_body)

        if "PHIEU GIAO HANG" not in text_body:
            continue

        idx = text_body.find("PHIEU GIAO HANG")
        sub = text_body[idx: idx + 200]

        sm, prso, date = extract_line(sub)

        if (sm or prso) and date:
            return sm, prso, date

    return None, None, None

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file):

    results = []
    images = convert_from_bytes(file.read(), dpi=150)

    for i, img in enumerate(images, 1):
        sm, prso, date = ocr_extract(img)

        if sm or prso:
            results.append({
                "SM": sm or "",
                "PR/SO": prso or "",
                "Ngày": date or "",
                "Trang": i
            })

    return results

# =========================
# BUTTON
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý"):
        st.session_state.run = True

# =========================
# RUN PROCESS
# =========================
if st.session_state.run and uploaded_files:

    st.write("⏳ Đang xử lý...")

    tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

    with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

        has_data = False

        for f in uploaded_files:

            data = extract_pdf(f)

            if data:
                has_data = True

                df = pd.DataFrame(data)
                df.insert(0, "STT", range(1, len(df)+1))

                name = os.path.splitext(f.name)[0][:31]
                df.to_excel(writer, sheet_name=name, index=False)

        if not has_data:
            df = pd.DataFrame([{"Thông báo": "Không có dữ liệu hợp lệ"}])
            df.to_excel(writer, sheet_name="KET_QUA", index=False)

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

    with open(tmp_excel.name, "rb") as f:
        st.download_button("📥 Tải Excel", f, file_name="result.xlsx")

    st.success("✅ Xử lý xong!")

    st.session_state.run = False
