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
for key in ["processing","done","clear_uploader","last_uploaded_names","excel_file"]:
    if key not in st.session_state:
        st.session_state[key] = False if key!="last_uploaded_names" else []

# =========================
# HEADER
# =========================
st.title("🚀 THL PDF → EXCEL")

# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# OCR (TỐI GIẢN - NHANH)
# =========================
def ocr_extract(img):

    # 👉 chỉ lấy vùng trên
    w, h = img.size
    img_crop = img.crop((0, 0, w, int(h * 0.4)))

    def read(image):
        text = pytesseract.image_to_string(
            image,
            lang='eng',
            config='--oem 3 --psm 6'
        )
        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
        return sm, date

    # 👉 CHỈ 2 HƯỚNG
    for v in [img_crop, img_crop.rotate(180, expand=True)]:
        sm, date = read(v)
        if sm and date:
            return sm.group(1), date.group(1)

    return None, None

# =========================
# PROCESS FILE
# =========================
def process_file(file, progress_bar):

    results = []

    # 👉 CHỈ convert 1 lần (DPI thấp)
    images = convert_from_bytes(file.read(), dpi=90)

    total = len(images)

    for i, img in enumerate(images):

        # 👉 update UI mỗi 3 trang (giảm lag)
        if i % 3 == 0:
            progress_bar.progress(int((i+1)/total*100))

        sm, date = ocr_extract(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date,
                "Trang": i+1
            })

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    if st.button("🚀 Bắt đầu xử lý"):

        start = time.time()

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for f in uploaded_files:

                st.write(f"📄 Đang xử lý: {f.name}")
                progress_bar = st.progress(0)

                data = process_file(f, progress_bar)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # =========================
        # FORMAT EXCEL
        # =========================
        wb = load_workbook(tmp_excel.name)

        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:

            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 2

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border

            for cell in ws[1]:
                cell.font = Font(bold=True)

        wb.save(tmp_excel.name)

        # =========================
        # DOWNLOAD
        # =========================
        with open(tmp_excel.name, "rb") as f:
            data = f.read()

        b64 = base64.b64encode(data).decode()

        st.success(f"🎉 Xong trong {round(time.time()-start,2)}s")

        st.markdown(f"""
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="result.xlsx">
        📥 Tải file Excel
        </a>
        """, unsafe_allow_html=True)
