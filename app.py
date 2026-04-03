import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import zipfile
import os
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="OCR PDF Tool", layout="wide")

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #0ea5e9, #22c55e);
}

h1 {
    text-align: center;
    color: white !important;
    font-weight: 900;
}

.block {
    background: rgba(255,255,255,0.92);
    padding: 12px;
    border-radius: 12px;
    font-weight: 600;
    box-shadow: 0 6px 20px rgba(0,0,0,0.15);
    text-align: center;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

</style>
""", unsafe_allow_html=True)

st.title("📄 OCR MULTI PDF DASHBOARD (GRID MODE)")


# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')

    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)

    return (sm.group(1), date.group(1)) if sm and date else (None, None)


# =========================
# PROCESS FILE
# =========================
def extract_pdf(file, status_box, file_idx, total_files):
    results = []

    images = convert_from_bytes(file.read(), dpi=150)
    total_pages = len(images)

    for i, img in enumerate(images, start=1):

        # update đúng 1 ô (không xuống dòng)
        status_box.markdown(
            f"""
            <div class="block">
            📁 {file.name}<br>
            📄 {i}/{total_pages} | ⚡ Processing
            </div>
            """,
            unsafe_allow_html=True
        )

        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)

        if sm and date:
            results.append({"SM": sm, "Ngày": date})

    status_box.markdown(
        f"""
        <div class="block">
        📁 {file.name}<br>
        ✅ DONE
        </div>
        """,
        unsafe_allow_html=True
    )

    return results


# =========================
# UPLOAD
# =========================
uploaded_files = st.file_uploader(
    "📤 Upload nhiều PDF",
    type=["pdf"],
    accept_multiple_files=True
)


# =========================
# RUN
# =========================
if uploaded_files:

    st.success(f"Đã chọn {len(uploaded_files)} file")

    # 🔥 CHIA CỘT THEO SỐ FILE
    cols = st.columns(len(uploaded_files))

    status_boxes = []

    for i in range(len(uploaded_files)):
        with cols[i]:
            box = st.empty()
            status_boxes.append(box)

    if st.button("🚀 START OCR"):

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:

            for idx, file in enumerate(uploaded_files):

                data = extract_pdf(
                    file,
                    status_boxes[idx],
                    idx,
                    len(uploaded_files)
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))

                    base = os.path.splitext(file.name)[0]
                    excel_name = f"{base}.xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name, index=False)
                        wb = load_workbook(tmp.name)
                        ws = wb.active

                        for col in ws.columns:
                            max_len = 0
                            col_letter = col[0].column_letter
                            for c in col:
                                if c.value:
                                    max_len = max(max_len, len(str(c.value)))
                            ws.column_dimensions[col_letter].width = max_len + 3

                        wb.save(tmp.name)
                        zipf.write(tmp.name, excel_name)

        st.success("🎉 HOÀN TẤT!")

        with open(zip_buffer.name, "rb") as f:
            st.download_button(
                "📥 DOWNLOAD ZIP",
                f,
                file_name="ocr_results.zip"
            )
