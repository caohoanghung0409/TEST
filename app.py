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
[data-testid="stFileUploader"] { border: 2px dashed #93c5fd; padding: 25px; border-radius: 18px; background: white; }
div.stButton > button { background: linear-gradient(135deg,#3b82f6,#22c55e); color:white; border:none; border-radius:12px; padding:12px 24px; font-weight:600; width: 100%; }
.new-btn button { background: linear-gradient(135deg,#f59e0b,#ef4444) !important; }
.file-row { margin-top:12px; padding:10px; border-radius:12px; background:white; box-shadow:0 2px 8px rgba(0,0,0,0.05); }
.progress { height:8px; background:#e5e7eb; border-radius:999px; overflow:hidden; margin-top:6px; }
.progress-bar { height:100%; background:linear-gradient(90deg,#3b82f6,#22c55e); transition: width 0.3s ease; }
.global-bar { position:relative; height:20px; background:#e5e7eb; border-radius:999px; overflow:hidden; margin-top:10px;}
.global-fill { height:100%; transition: width 0.4s ease; }
.global-text { position:absolute; width:100%; text-align:center; font-size:12px; font-weight:700; top:0; line-height:20px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="header">🚀 THL PDF → EXCEL </div>', unsafe_allow_html=True)

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"
uploaded_files = st.file_uploader("📂 Chọn file PDF", type=["pdf"], accept_multiple_files=True, key=uploader_key)

if uploaded_files:
    current_names = [f.name for f in uploaded_files]
    if current_names != st.session_state.last_uploaded_names:
        st.session_state.processing = False
        st.session_state.done = False
        st.session_state.last_uploaded_names = current_names

# =========================
# OCR OPTIMIZED
# =========================
def ocr_extract(img):
    def read(image):
        text = pytesseract.image_to_string(image, lang='eng', config='--oem 3 --psm 6')
        sm = re.search(r"(SM\d{4}\.\d{4})", text)
        date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
        return sm, date

    w, h = img.size
    # Chỉ thử 2 biến thể quan trọng nhất để tăng tốc
    for variant in [img, img.crop((0, 0, w, int(h * 0.4)))]:
        sm, date = read(variant)
        if sm and date:
            return sm.group(1), date.group(1)
    return None, None

# =========================
# MAIN LOGIC
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
        
        # 1. Đếm tổng số trang cực nhanh (DPI thấp)
        total_pages_all = 0
        file_data_list = []
        for f in uploaded_files:
            bytes_data = f.read()
            file_data_list.append(bytes_data)
            # Dùng DPI=10 chỉ để đếm trang, không tốn tài nguyên
            temp_imgs = convert_from_bytes(bytes_data, dpi=10)
            total_pages_all += len(temp_imgs)
        
        processed_pages = 0
        tmp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        with pd.ExcelWriter(tmp_excel.name, engine='openpyxl') as writer:
            for idx, bytes_data in enumerate(file_data_list):
                results = []
                # Chuyển đổi thật để OCR (DPI 120 là mức cân bằng tốc độ)
                images = convert_from_bytes(bytes_data, dpi=120)
                total_file_pages = len(images)

                for i, img in enumerate(images, start=1):
                    processed_pages += 1
                    
                    # Cập nhật Global Bar
                    g_percent = int((processed_pages / total_pages_all) * 100)
                    elapsed = time.time() - start_time
                    speed = processed_pages / elapsed if elapsed > 0 else 0
                    eta = int((total_pages_all - processed_pages) / speed) if speed > 0 else 0
                    
                    global_box.markdown(f"""
                        <div class="global-bar">
                            <div class="global-fill" style="width:{g_percent}%; background:#22c55e;"></div>
                            <div class="global-text">Tổng tiến độ: {g_percent}% | ETA: {eta//60}m {eta%60}s</div>
                        </div>
                    """, unsafe_allow_html=True)

                    # Cập nhật File Bar
                    f_percent = int((i / total_file_pages) * 100)
                    boxes[idx].markdown(f"""
                        <div class="file-row">
                            📄 {uploaded_files[idx].name} ({i}/{total_file_pages})
                            <div class="progress"><div class="progress-bar" style="width:{f_percent}%"></div></div>
                        </div>
                    """, unsafe_allow_html=True)

                    sm, date = ocr_extract(img)
                    if sm and date:
                        results.append({"SM": sm, "Ngày": date, "Trang": i})

                if results:
                    df = pd.DataFrame(results)
                    df.insert(0, "STT", range(1, len(df)+1))
                    sheet_name = os.path.splitext(uploaded_files[idx].name)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Định dạng Excel
        wb = load_workbook(tmp_excel.name)
        for ws in wb.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 15
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        wb.save(tmp_excel.name)

        st.session_state.excel_file = tmp_excel.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

if st.session_state.done:
    st.success("🎉 Xử lý hoàn tất!")
    with open(st.session_state.excel_file, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    
    st.markdown(f'<iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" style="display:none;"></iframe>', unsafe_allow_html=True)
    
    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
