import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
from pypdf import PdfReader
import pandas as pd
import re
import tempfile
import os
import time
import base64
import io
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
# STYLE (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header {
    font-size:22px;
    font-weight:700;
    margin-bottom:10px;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #93c5fd;
    padding: 25px;
    border-radius: 18px;
    background: white;
    transition: 0.3s;
}
[data-testid="stFileUploader"]:hover {
    border-color:#3b82f6;
}

div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:12px 24px;
    font-weight:600;
    font-size:15px;
    box-shadow:0 4px 14px rgba(0,0,0,0.15);
    transition: all 0.25s ease;
}
div.stButton > button:hover {
    transform: translateY(-2px) scale(1.02);
}

.new-btn button {
    background: linear-gradient(135deg,#f59e0b,#ef4444) !important;
}

.process-btn {
    margin-top: 25px;
    margin-bottom: 15px;
}

.file-row {
    margin-top:12px;
    padding:10px;
    border-radius:12px;
    background:white;
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}

.progress {
    height:8px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
    margin-top:6px;
}
.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#3b82f6,#22c55e);
    transition: width 0.3s ease;
}

.global-wrap { margin:15px 0; }

.global-bar {
    position:relative;
    height:20px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
}

.global-fill {
    height:100%;
    border-radius:999px;
    transition: width 0.4s ease;
}

.global-fill::before {
    content:"";
    position:absolute;
    width:100%;
    height:100%;
    background: repeating-linear-gradient(
        45deg,
        rgba(255,255,255,0.2) 0,
        rgba(255,255,255,0.2) 10px,
        transparent 10px,
        transparent 20px
    );
    animation: move 1s linear infinite;
}

@keyframes move {
    from { background-position: 0 0; }
    to { background-position: 40px 0; }
}

.global-text {
    position:absolute;
    width:100%;
    text-align:center;
    font-size:12px;
    font-weight:700;
    top:0;
    line-height:20px;
}

.global-meta {
    display:flex;
    justify-content:space-between;
    font-size:13px;
    margin-bottom:6px;
}

.loading {
    font-size:14px;
    color:#475569;
    margin-top:10px;
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
# HÀM BÓC TÁCH NHANH (REGEX & OCR)
# =========================
REGEX_SM = re.compile(r"(SM\d{4}\.\d{4})")
REGEX_DATE = re.compile(r"(\d{2}/\d{2}/\d{4})")

# Giới hạn ký tự và chế độ nhận diện dòng đơn lẻ
TESS_CONFIG = '--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789./SM -c load_system_dawg=0 -c load_freq_dawg=0'

def parse_text(text):
    if not text:
        return None, None
    sm = REGEX_SM.search(text)
    date = REGEX_DATE.search(text)
    return (sm.group(1) if sm else None), (date.group(1) if date else None)

def fast_ocr_zone(img_zone):
    # Chuyển sang ảnh xám và tăng độ tương phản để Tesseract đọc cực nhanh
    gray = img_zone.convert('L')
    # Nhị phân hóa (chữ đen, nền trắng hoàn toàn)
    binary = gray.point(lambda p: 255 if p > 165 else 0)
    text = pytesseract.image_to_string(binary, lang='eng', config=TESS_CONFIG)
    return parse_text(text)

def ocr_extract_optimized(img):
    w, h = img.size

    # 1. Cắt riêng góc trên bên phải (Top-Right: nơi in SM và Ngày của phiếu giao hàng)
    # Vùng này chỉ bằng 1/6 diện tích trang giấy -> OCR chạy cực kỳ nhanh (mất ~0.1s)
    top_right = img.crop((int(w * 0.4), 0, w, int(h * 0.35)))
    sm, date = fast_ocr_zone(top_right)
    if sm and date:
        return sm, date

    # 2. Nếu thiếu, mở rộng quét toàn bộ nửa trên trang
    top_half = img.crop((0, 0, w, int(h * 0.35)))
    sm, date = fast_ocr_zone(top_half)
    if sm and date:
        return sm, date

    # 3. Chỉ khi giấy bị scan ngược 180 độ: xoay và cắt đúng vùng góc đầu
    img_180 = img.rotate(180, expand=True)
    top_right_180 = img_180.crop((int(w * 0.4), 0, w, int(h * 0.35)))
    sm, date = fast_ocr_zone(top_right_180)
    if sm and date:
        return sm, date

    return None, None

# =========================
# GLOBAL BAR
# =========================
def render_global_bar(percent, eta):
    eta_text = "Sắp xong..." if eta <= 0 else f"{eta//60}m {eta%60}s"

    return f"""
<div class="global-wrap">
    <div class="global-meta">
        <div>⚡ {percent}%</div>
        <div>⏳ {eta_text}</div>
    </div>
    <div class="global-bar">
        <div class="global-fill" style="width:{percent}%; background:linear-gradient(90deg,#3b82f6,#22c55e);"></div>
        <div class="global-text">{percent}%</div>
    </div>
</div>
"""

# =========================
# PROCESS CHÍNH
# =========================
def extract_pdf(file_bytes, file_name, box, global_box, start_time, processed_pages, total_pages_all):
    results = []

    # BƯỚC 1: Thử đọc trực tiếp Text gốc từ PDF bằng PyPDF (mất 0.01s/trang)
    pypdf_reader = None
    try:
        pypdf_reader = PdfReader(io.BytesIO(file_bytes))
        total_pages = len(pypdf_reader.pages)
    except Exception:
        total_pages = 1

    # BƯỚC 2: Chuyển đổi PDF sang ảnh với DPI=100 (tối ưu tốc độ, vừa chuẩn nét để OCR chữ to)
    # Chỉ render ảnh khi cần
    images = convert_from_bytes(file_bytes, dpi=100)
    total_pages = len(images)

    for i in range(1, total_pages + 1):
        processed_pages[0] += 1

        percent = int((i / total_pages) * 100)
        global_percent = int((processed_pages[0] / total_pages_all) * 100)

        elapsed = time.time() - start_time
        speed = processed_pages[0] / elapsed if elapsed > 0 else 0
        remaining = total_pages_all - processed_pages[0]
        eta = int(remaining / speed) if speed > 0 else 0

        global_box.markdown(render_global_bar(global_percent, eta), unsafe_allow_html=True)
        box.markdown(f"""
<div class="file-row">
📄 {file_name} — Trang {i}/{total_pages} ({percent}%)
<div class="progress">
<div class="progress-bar" style="width:{percent}%"></div>
</div>
</div>
""", unsafe_allow_html=True)

        sm, date = None, None

        # Cách A: Thử trích xuất trực tiếp từ Text gốc (nếu PDF có text thì không tốn 1ms nào cho OCR)
        if pypdf_reader and (i - 1) < len(pypdf_reader.pages):
            try:
                raw_text = pypdf_reader.pages[i - 1].extract_text() or ""
                sm, date = parse_text(raw_text)
            except Exception:
                pass

        # Cách B: Nếu không có text gốc (ảnh scan), mới kích hoạt OCR vùng chọn
        if not (sm and date):
            img = images[i - 1]
            sm, date = ocr_extract_optimized(img)

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
        st.markdown('<div class="process-btn">', unsafe_allow_html=True)
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.processing:
        st.markdown('<div class="loading">⏳ Đang xử lý... vui lòng chờ</div>', unsafe_allow_html=True)

        start_time = time.time()

        file_data_list = []
        total_pages_all = 0

        # Lấy trước số trang
        for f in uploaded_files:
            b = f.read()
            file_data_list.append((f.name, b))
            try:
                reader = PdfReader(io.BytesIO(b))
                total_pages_all += len(reader.pages)
            except Exception:
                total_pages_all += 1

        total_pages_all = max(total_pages_all, 1)
        processed_pages = [0]

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:
            for i, (fname, fbytes) in enumerate(file_data_list):
                data = extract_pdf(
                    fbytes, fname, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))
                    sheet_name = os.path.splitext(fname)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Định dạng viền và cột Excel
        wb = load_workbook(tmp_excel.name)
        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 10)

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

    st.markdown('<div class="new-btn">', unsafe_allow_html=True)
    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
