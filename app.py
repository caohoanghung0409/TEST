import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import time
from io import BytesIO

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# TITLE
# =========================
st.markdown("## 🚀 THL PDF → EXCEL (WEB)")

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
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# EXTRACT PDF
# =========================
def extract_pdf(file, progress_bar, status_text):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):
        percent = int((i / total_pages) * 100)
        progress_bar.progress(i / total_pages)
        status_text.text(f"📄 {file.name} - Trang {i}/{total_pages} ({percent}%)")

        # crop vùng trên (tăng tốc + chính xác hơn)
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({
                "Trang": i,
                "SM": sm,
                "Ngày": date
            })

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    st.info(f"📦 Tổng số file: {len(uploaded_files)}")

    if st.button("🚀 Bắt đầu xử lý"):

        start_time = time.time()

        all_sheets = {}

        for file in uploaded_files:

            st.markdown(f"### 📄 Đang xử lý: {file.name}")

            progress_bar = st.progress(0)
            status_text = st.empty()

            data = extract_pdf(file, progress_bar, status_text)

            if data:
                df = pd.DataFrame(data)
                df.insert(0, "STT", range(1, len(df) + 1))

                # Excel giới hạn 31 ký tự sheet name
                sheet_name = file.name[:31]

                all_sheets[sheet_name] = df
            else:
                st.warning(f"⚠️ Không tìm thấy dữ liệu trong file {file.name}")

        # =========================
        # EXPORT EXCEL (RAM)
        # =========================
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)

        elapsed = time.time() - start_time

        # =========================
        # DONE
        # =========================
        st.success(f"🎉 HOÀN THÀNH sau {round(elapsed, 2)}s")

        # =========================
        # DOWNLOAD
        # =========================
        st.download_button(
            label="📥 Tải file Excel",
            data=output,
            file_name="ket_qua.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
