import streamlit as st
from pypdf import PdfReader
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import time
import base64
import io
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

# =========================================================
# TỰ ĐỘNG NHẬN DIỆN ĐƯỜNG DẪN TESSERACT VÀ POPPLER
# =========================================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
TESSERACT_EXE = os.path.join(CURRENT_DIR, "tesseract", "tesseract.exe")
POPPLER_BIN = os.path.join(CURRENT_DIR, "poppler", "Library", "bin")

if not os.path.exists(POPPLER_BIN):
    POPPLER_BIN = os.path.join(CURRENT_DIR, "poppler", "bin")

if os.path.exists(TESSERACT_EXE):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

# =========================================================
# CẤU HÌNH GIAO DIỆN
# =========================================================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

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

st.markdown('<div class="header">🚀 THL PDF → EXCEL (SIÊU TỐC & CHÍNH XÁC)</div>', unsafe_allow_html=True)

# =========================================================
# FILE UPLOADER
# =========================================================
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

# =========================================================
# BỘ LỌC THÔNG MINH TIỀN PHONG & REGEX
# =========================================================
REGEX_TIENPHONG = re.compile(
    r"(TIEN\s*PHONG|TIỀN\s*PHONG|NHUATIENPHONG|TIENPHONGNAM|PLASTIC|Đồng\s*An|DVKH|nhuatienphong|Acumatica|PHIẾU\s*GIAO\s*HÀNG|PHIEU\s*GIAO\s*HANG)",
    re.IGNORECASE
)
REGEX_SM = re.compile(r"SM\s*(\d{4})[\.\,\-\s\_]?(\d{4})", re.IGNORECASE)
REGEX_DATE = re.compile(r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})")

def analyze_text(text):
    if not text:
        return None, None, False

    if not REGEX_TIENPHONG.search(text):
        return None, None, False

    sm_match = REGEX_SM.search(text)
    if not sm_match:
        return None, None, True

    clean_sm = f"SM{sm_match.group(1)}.{sm_match.group(2)}"

    clean_date = ""
    date_context = re.search(r"Ng[àa]y[\s\:\.]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})", text, re.IGNORECASE)
    if date_context:
        clean_date = re.sub(r"[\-\.]", "/", date_context.group(1))
    else:
        date_match = REGEX_DATE.search(text)
        if date_match:
            clean_date = re.sub(r"[\-\.]", "/", date_match.group(1))

    return clean_sm, clean_date, True

# OCR dự phòng gọn nhẹ (chỉ chạy khi không có text số)
def ocr_fallback(img):
    w, h = img.size
    # 1. Quét nửa trên 45% (chiều thẳng)
    crop_top = img.crop((0, 0, w, int(h * 0.45)))
    text = pytesseract.image_to_string(crop_top, lang='eng', config='--oem 3 --psm 6')
    sm, date, is_tp = analyze_text(text)
    if is_tp and sm:
        return sm, date

    # 2. Quét trường hợp ngược 180 độ
    img_180 = img.rotate(180, expand=True)
    crop_bottom = img_180.crop((0, 0, w, int(h * 0.45)))
    text_180 = pytesseract.image_to_string(crop_bottom, lang='eng', config='--oem 3 --psm 6')
    sm, date, is_tp = analyze_text(text_180)
    if is_tp and sm:
        return sm, date

    return None, None

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

# =========================================================
# XỬ LÝ SIÊU TỐC TỪNG FILE PDF
# =========================================================
def extract_pdf_fast(file_bytes, file_name, box, global_box, start_time, processed_pages, total_pages_all):
    results = []
    reader = PdfReader(io.BytesIO(file_bytes))
    total_pages = len(reader.pages)
    poppler_dir = POPPLER_BIN if os.path.exists(POPPLER_BIN) else None

    for i, page in enumerate(reader.pages, start=1):
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

        # 1. ĐỌC TEXT TRỰC TIẾP (Siêu nhanh 0.005s, bắt trọn vẹn trang 23)
        raw_text = page.extract_text() or ""
        sm, date, is_tp = analyze_text(raw_text)

        # 2. DỰ PHÒNG OCR (Chỉ kích hoạt nếu trang là ảnh scan thuần)
        if not (is_tp and sm):
            try:
                if poppler_dir:
                    imgs = convert_from_bytes(file_bytes, first_page=i, last_page=i, dpi=130, poppler_path=poppler_dir)
                else:
                    imgs = convert_from_bytes(file_bytes, first_page=i, last_page=i, dpi=130)
                if imgs:
                    sm_ocr, date_ocr = ocr_fallback(imgs[0])
                    if sm_ocr:
                        sm, date = sm_ocr, date_ocr
            except Exception:
                pass

        if sm:
            results.append({
                "SM": sm,
                "Ngày": date if date else "",
                "Trang": i
            })

    return results

# =========================================================
# LUỒNG XỬ LÝ CHÍNH
# =========================================================
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
        st.markdown('<div class="loading">⏳ Đang xử lý siêu tốc... vui lòng chờ</div>', unsafe_allow_html=True)
        start_time = time.time()

        file_data_list = []
        total_pages_all = 0
        for f in uploaded_files:
            b = f.read()
            file_data_list.append((f.name, b))
            try:
                r = PdfReader(io.BytesIO(b))
                total_pages_all += len(r.pages)
            except Exception:
                total_pages_all += 1

        total_pages_all = max(total_pages_all, 1)
        processed_pages = [0]

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:
            for i, (fname, fbytes) in enumerate(file_data_list):
                data = extract_pdf_fast(
                    fbytes, fname, boxes[i], global_box,
                    start_time, processed_pages, total_pages_all
                )

                sheet_name = os.path.splitext(fname)[0][:31]
                if data:
                    df = pd.DataFrame(data)
                    df.insert(0, "STT", range(1, len(df) + 1))
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    df_empty = pd.DataFrame([{"Thông báo": "Không tìm thấy mã SM Tiền Phong hợp lệ"}])
                    df_empty.to_excel(writer, sheet_name=sheet_name, index=False)

        # Định dạng kẻ bảng Excel
        wb = load_workbook(tmp_excel.name)
        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 12)

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

# =========================================================
# TẢI XUỐNG VÀ XỬ LÝ LẠI
# =========================================================
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
