import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import time
import base64
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

# =========================
# CONFIG
# =========================
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
    border: 2px dashed #93c5fd; padding: 25px; border-radius: 18px; background: white; transition: 0.3s;
}
[data-testid="stFileUploader"]:hover { border-color:#3b82f6; }
div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e); color:white; border:none;
    border-radius:12px; padding:12px 24px; font-weight:600; font-size:15px;
    box-shadow:0 4px 14px rgba(0,0,0,0.15); transition: all 0.25s ease;
}
div.stButton > button:hover { transform: translateY(-2px) scale(1.02); }
.new-btn button { background: linear-gradient(135deg,#f59e0b,#ef4444) !important; }
.process-btn { margin-top: 25px; margin-bottom: 15px; }
.file-row { margin-top:12px; padding:10px; border-radius:12px; background:white; box-shadow:0 2px 8px rgba(0,0,0,0.05); }
.progress { height:8px; background:#e5e7eb; border-radius:999px; overflow:hidden; margin-top:6px; }
.progress-bar { height:100%; background:linear-gradient(90deg,#3b82f6,#22c55e); transition: width 0.3s ease; }
.global-wrap { margin:15px 0; }
.global-bar { position:relative; height:20px; background:#e5e7eb; border-radius:999px; overflow:hidden; }
.global-fill { height:100%; border-radius:999px; transition: width 0.4s ease; }
.global-text { position:absolute; width:100%; text-align:center; font-size:12px; font-weight:700; top:0; line-height:20px; }
.global-meta { display:flex; justify-content:space-between; font-size:13px; margin-bottom:6px; }
.loading { font-size:14px; color:#475569; margin-top:10px; }
</style>
""", unsafe_allow_html=True)

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
# THUẬT TOÁN OCR THÔNG MINH (CHÍNH XÁC + TỐI ƯU SỐ LẦN QUÉT)
# =========================
REGEX_SM = re.compile(r"(SM\d{4}\.\d{4})")
REGEX_DATE = re.compile(r"(\d{2}/\d{2}/\d{4})")

def extract_patterns(txt):
    sm = REGEX_SM.search(txt)
    date = REGEX_DATE.search(txt)
    return (sm.group(1) if sm else None), (date.group(1) if date else None)

def ocr_smart(img):
    w, h = img.size
    
    # BƯỚC 1: Quét vùng 40% trên cùng của trang gốc (Nhanh nhất: chiếm 80% trường hợp)
    top_crop = img.crop((0, 0, w, int(h * 0.42)))
    txt1 = pytesseract.image_to_string(top_crop, lang='eng', config='--oem 3 --psm 6')
    sm, date = extract_patterns(txt1)
    if sm and date:
        return sm, date

    # BƯỚC 2: Quét toàn bộ ảnh gốc (dành cho phiếu có layout đặc thù)
    txt2 = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm, date = extract_patterns(txt2)
    if sm and date:
        return sm, date

    # BƯỚC 3: Nếu ảnh bị scan lộn ngược (180 độ - trường hợp lỗi phổ biến nhất)
    img_180 = img.rotate(180, expand=True)
    top_180 = img_180.crop((0, 0, w, int(h * 0.42)))
    txt3 = pytesseract.image_to_string(top_180, lang='eng', config='--oem 3 --psm 6')
    sm, date = extract_patterns(txt3)
    if sm and date:
        return sm, date

    # BƯỚC 4: Dự phòng cuối cùng cho ảnh xoay ngang (90, 270)
    for rot in (90, 270):
        img_r = img.rotate(rot, expand=True)
        txt_r = pytesseract.image_to_string(img_r, lang='eng', config='--oem 3 --psm 6')
        sm, date = extract_patterns(txt_r)
        if sm and date:
            return sm, date

    return None, None

# =========================
# TIẾN TRÌNH UI
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
# XỬ LÝ CHÍNH
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

        # 1. Chuyển đổi PDF sang ảnh (chỉ chạy 1 lần duy nhất với DPI 140 tối ưu)
        files_data = []
        total_pages_all = 0
        for f in uploaded_files:
            file_bytes = f.read()
            imgs = convert_from_bytes(file_bytes, dpi=140)
            files_data.append((f.name, imgs))
            total_pages_all += len(imgs)

        total_pages_all = max(total_pages_all, 1)
        processed_pages = 0

        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:
            for file_idx, (fname, images) in enumerate(files_data):
                total_pages = len(images)
                results = []

                for i, img in enumerate(images, start=1):
                    processed_pages += 1
                    
                    percent = int((i / total_pages) * 100)
                    global_percent = int((processed_pages / total_pages_all) * 100)
                    
                    elapsed = time.time() - start_time
                    speed = processed_pages / elapsed if elapsed > 0 else 0
                    remaining = total_pages_all - processed_pages
                    eta = int(remaining / speed) if speed > 0 else 0

                    global_box.markdown(render_global_bar(global_percent, eta), unsafe_allow_html=True)
                    boxes[file_idx].markdown(f"""
<div class="file-row">
📄 {fname} — Trang {i}/{total_pages} ({percent}%)
<div class="progress">
<div class="progress-bar" style="width:{percent}%"></div>
</div>
</div>
""", unsafe_allow_html=True)

                    sm, date = ocr_smart(img)
                    if sm and date:
                        results.append({
                            "SM": sm,
                            "Ngày": date,
                            "Trang": i
                        })

                if results:
                    df = pd.DataFrame(results)
                    df.insert(0, "STT", range(1, len(df) + 1))
                    sheet_name = os.path.splitext(fname)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Định dạng thẩm mỹ Excel
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
# TẢI XUỐNG
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
