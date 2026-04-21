import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import time
import base64
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

st.markdown("## 🚀 THL PDF → EXCEL (SCAN FAST VERSION)")

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF scan A4",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# OCR OPTIMIZED
# =========================
def ocr_extract(img):

    def read(image):
        text = pytesseract.image_to_string(
            image,
            lang='eng',
            config='--oem 3 --psm 6'
        )

        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

        return sm, date

    w, h = img.size

    # 🚀 CHỈ LẤY VÙNG TRÊN (SM + DATE thường nằm đây)
    top = img.crop((0, 0, w, int(h * 0.45)))

    # ⚡ 3 VARIANT LÀ ĐỦ (cân bằng tốc độ + chính xác)
    variants = [
        top,
        top.rotate(180, expand=True),
        img  # fallback cuối
    ]

    for v in variants:
        sm, date = read(v)
        if sm and date:
            return sm.group(1), date.group(1)

    return None, None

# =========================
# PROCESS PDF
# =========================
def process_pdf(file):

    results = []

    # ⚡ DPI tối ưu cho scan
    images = convert_from_bytes(file.read(), dpi=110)

    for i, img in enumerate(images, start=1):

        sm, date = ocr_extract(img)

        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date,
                "Trang": i
            })

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    if st.button("🚀 BẮT ĐẦU XỬ LÝ"):

        start = time.time()

        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        has_data = False

        with pd.ExcelWriter(tmp_file.name, engine='openpyxl') as writer:

            for f in uploaded_files:

                data = process_pdf(f)

                if data:
                    has_data = True

                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = f.name[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 🛟 luôn đảm bảo có sheet
            if not has_data:
                pd.DataFrame([{
                    "Thông báo": "Không có dữ liệu hợp lệ"
                }]).to_excel(writer, sheet_name="NO_DATA", index=False)

        # ================= FORMAT EXCEL =================
        wb = load_workbook(tmp_file.name)

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

        wb.save(tmp_file.name)

        st.success(f"🎉 XONG TRONG {round(time.time()-start,2)} GIÂY")

        with open(tmp_file.name, "rb") as f:
            st.download_button(
                "📥 DOWNLOAD EXCEL",
                f,
                file_name="THL_RESULT.xlsx"
            )
