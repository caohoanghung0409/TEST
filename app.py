import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
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
for key in ["processing", "done", "clear_uploader", "last_uploaded_names", "excel_file"]:
    if key not in st.session_state:
        st.session_state[key] = False if key in ["processing", "done"] else []

# =========================
# REGEX (TỐI ƯU)
# =========================
SM_REGEX = re.compile(r"(SM\d{4}\.\d{4})")
DATE_REGEX = re.compile(r"(\d{2}/\d{2}/\d{4})")

# =========================
# STYLE (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }
.header { font-size:22px; font-weight:700; margin-bottom:10px; }
div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:12px 24px;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="header">🚀 THL PDF → EXCEL</div>', unsafe_allow_html=True)

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

current_names = [f.name for f in uploaded_files] if uploaded_files else []

if current_names != st.session_state.last_uploaded_names:
    st.session_state.processing = False
    st.session_state.done = False
    st.session_state.last_uploaded_names = current_names

# =========================
# OCR FUNCTION (NHANH + ỔN ĐỊNH)
# =========================
def ocr_extract(img):

    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')

    sm = SM_REGEX.search(text)
    date = DATE_REGEX.search(text)

    if sm and date:
        return sm.group(1), date.group(1)

    # thử thêm 1 hướng xoay (nhẹ thôi để tăng tốc)
    img2 = img.rotate(180, expand=True)
    text2 = pytesseract.image_to_string(img2, lang='eng', config='--oem 3 --psm 6')

    sm = SM_REGEX.search(text2)
    date = DATE_REGEX.search(text2)

    if sm and date:
        return sm.group(1), date.group(1)

    return None, None

# =========================
# PROGRESS BAR
# =========================
def render_global_bar(percent, eta):
    eta_text = "Sắp xong..." if eta == 0 else f"{eta//60}m {eta%60}s"
    return f"""
    <div>
        <div>⚡ {percent}% | ⏳ {eta_text}</div>
        <div style="height:16px;background:#ddd;border-radius:10px;">
            <div style="width:{percent}%;height:100%;
            background:linear-gradient(90deg,#3b82f6,#22c55e);"></div>
        </div>
    </div>
    """

# =========================
# EXTRACT PDF
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages_all):

    results = []

    file_bytes = file.getvalue()

    # lấy số trang nhanh
    total_pages = int(pdfinfo_from_bytes(file_bytes)["Pages"])

    # giảm DPI để tăng tốc
    images = convert_from_bytes(file_bytes, dpi=110)

    for i, img in enumerate(images, start=1):

        processed_pages[0] += 1

        percent = int((processed_pages[0] / total_pages_all) * 100)

        elapsed = time.time() - start_time
        speed = processed_pages[0] / elapsed if elapsed > 0 else 0
        remaining = total_pages_all - processed_pages[0]
        eta = int(remaining / speed) if speed > 0 else 0

        global_box.markdown(render_global_bar(percent, eta), unsafe_allow_html=True)

        box.markdown(f"📄 {file.name} — Trang {i}/{total_pages}")

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

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start_time = time.time()

        file_data = []
        total_pages_all = 0

        # preload bytes + pages
        for f in uploaded_files:
            b = f.getvalue()
            pages = int(pdfinfo_from_bytes(b)["Pages"])
            file_data.append((f, b))
            total_pages_all += pages

        processed_pages = [0]

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            has_data = False

            for i, (f, b) in enumerate(file_data):

                fake_file = type("F", (), {
                    "name": f.name,
                    "getvalue": lambda b=b: b
                })()

                data = extract_pdf(
                    fake_file,
                    boxes[i],
                    global_box,
                    start_time,
                    processed_pages,
                    total_pages_all
                )

                if data and len(data) > 0:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))
                    df.to_excel(writer, sheet_name=f.name[:31], index=False)
                    has_data = True

            # =========================
            # FIX CRASH OPENPYXL
            # =========================
            if not has_data:
                pd.DataFrame([{"THÔNG BÁO": "KHÔNG TRÍCH XUẤT ĐƯỢC DỮ LIỆU"}]).to_excel(
                    writer,
                    sheet_name="EMPTY",
                    index=False
                )

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
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 HOÀN THÀNH !!!")

    with open(st.session_state.excel_file, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    st.markdown(f"""
    <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
