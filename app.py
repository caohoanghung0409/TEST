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
if "processing" not in st.session_state:
    st.session_state.processing = False
if "done" not in st.session_state:
    st.session_state.done = False
if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False
if "last_uploaded_names" not in st.session_state:
    st.session_state.last_uploaded_names = []
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header { font-size:22px; font-weight:700; margin-bottom:10px; }

[data-testid="stFileUploader"] {
    border: 2px dashed #93c5fd;
    padding: 25px;
    border-radius: 18px;
    background: white;
}

div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:12px 24px;
    font-weight:600;
}

.new-btn button {
    background: linear-gradient(135deg,#f59e0b,#ef4444) !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 THL PDF → EXCEL </div>', unsafe_allow_html=True)

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
# OCR NÂNG CAO
# =========================
def clean_text(text):
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text

def extract_so_ngay(text):
    text = clean_text(text)

    # Bắt "Số" linh hoạt (Số, So, S0, S6)
    so_pattern = r"S[oố0O6]\s*[:\-]?\s*([A-Z0-9./\s]+)"
    so_match = re.search(so_pattern, text)

    so = None
    if so_match:
        so = so_match.group(1)
        so = so.strip()
        so = so.replace(" ", "")

        # Cắt nếu OCR ăn quá dài
        so = re.split(r"(Ngày|Ngay|\d{2}/\d{2}/\d{4})", so)[0]

    # Ngày
    date_match = re.search(r"\d{2}/\d{2}/\d{4}", text)
    date = date_match.group(0) if date_match else None

    return so, date

def ocr_extract(img):
    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )
    return extract_so_ngay(text)

# =========================
# GLOBAL BAR
# =========================
def render_global_bar(percent, eta):
    eta_text = "Sắp xong..." if eta == 0 else f"{eta//60}m {eta%60}s"
    return f"""
    <div>
        <b>{percent}%</b> | ⏳ {eta_text}
        <div style="background:#ddd;height:10px;border-radius:10px;">
            <div style="width:{percent}%;background:#22c55e;height:10px;border-radius:10px;"></div>
        </div>
    </div>
    """

# =========================
# EXTRACT PDF
# =========================
def extract_pdf(file, box, global_box, start_time, processed_pages, total_pages_all):

    results = []

    pdf_bytes = file.read()
    info = pdfinfo_from_bytes(pdf_bytes)
    total_pages = info["Pages"]

    for i in range(1, total_pages + 1):

        images = convert_from_bytes(
            pdf_bytes,
            dpi=130,
            first_page=i,
            last_page=i
        )

        img = images[0]

        processed_pages[0] += 1

        percent = int((i / total_pages) * 100)
        global_percent = int((processed_pages[0] / total_pages_all) * 100)

        elapsed = time.time() - start_time
        speed = processed_pages[0] / elapsed if elapsed > 0 else 0
        remaining = total_pages_all - processed_pages[0]
        eta = int(remaining / speed) if speed > 0 else 0

        global_box.markdown(render_global_bar(global_percent, eta), unsafe_allow_html=True)

        box.markdown(f"📄 {file.name} - Trang {i}/{total_pages} ({percent}%)")

        so, date = ocr_extract(img)

        if so and date:
            results.append({
                "Số": so,
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

        total_pages_all = 0
        for f in uploaded_files:
            info = pdfinfo_from_bytes(f.read())
            total_pages_all += info["Pages"]
            f.seek(0)

        processed_pages = [0]

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:

            for i, f in enumerate(uploaded_files):

                data = extract_pdf(
                    f, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df)+1))

                    sheet_name = os.path.splitext(f.name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

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
        data = f.read()

    b64 = base64.b64encode(data).decode()

    st.markdown(f"""
    <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>
    """, unsafe_allow_html=True)

    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
