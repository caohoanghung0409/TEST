"""
ỨNG DỤNG WEB TRÍCH XUẤT SỐ SM & NGÀY (NHỰA TIỀN PHONG)
- Tên file Excel cố định: PDF-TO-EXCEL.xlsx
- Tự động download ngay khi quét xong (đã sửa triệt để lỗi không tự tải).
- Đếm ngược thời gian (phút & giây).
- Đổi vị trí: File ở đầu, Thời gian ở cuối.
- Giao diện siêu gọn 1 màn hình.
"""

import os
import sys
import re
import io
import time
import base64
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pypdfium2 as pdfium
import pytesseract
from PIL import Image
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ==================== CẤU HÌNH TESSERACT OCR ====================
def setup_tesseract():
    if os.name == 'nt':
        default_paths = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tesseract-OCR", "tesseract.exe"),
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
        ]
        for path in default_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return

setup_tesseract()

# ==================== HÀM TIỆN ÍCH ====================
def format_time_str(total_seconds):
    """Định dạng thời gian ra phút và giây rõ ràng"""
    sec = max(int(total_seconds), 0)
    mins = sec // 60
    secs = sec % 60
    if mins > 0:
        return f"{mins} phút {secs:02d} giây"
    else:
        return f"0 phút {secs:02d} giây"

def sanitize_sheet_name(name, existing_names):
    """Quy chuẩn tên Sheet trong Excel (tối đa 31 ký tự, loại bỏ ký tự cấm)"""
    if name.lower().endswith('.pdf'):
        name = name[:-4]
    clean_name = re.sub(r'[\/\\\?\*\[\]\:]', '_', name).strip()
    if not clean_name:
        clean_name = "Sheet"
    base_name = clean_name[:31]
    final_name = base_name
    counter = 1
    while final_name.lower() in existing_names:
        suffix = f"_{counter}"
        max_len = 31 - len(suffix)
        final_name = base_name[:max_len] + suffix
        counter += 1
    existing_names.add(final_name.lower())
    return final_name

def extract_from_single_pdf(file_bytes, page_callback=None):
    """Trích xuất SM và Ngày từ 1 file PDF"""
    pdf = pdfium.PdfDocument(file_bytes)
    total_pages = len(pdf)
    records = []
    
    for page_idx in range(total_pages):
        page_num = page_idx + 1
        page = pdf[page_idx]
        bitmap = page.render(scale=2.0)
        img = bitmap.to_pil()
        w, h = img.size
        
        # Cắt 30% đầu trang
        crop = img.crop((0, 0, w, int(h * 0.30)))
        try:
            text = pytesseract.image_to_string(crop, lang='eng')
        except Exception:
            text = ""
            
        text_u = text.upper()
        is_tp = ('TIEN PHONG' in text_u or 'THIEU NIEN' in text_u or 'NHUATIENPHONG' in text_u)
        
        if is_tp:
            sm_match = re.search(r'[\$S§s]M[\s\.:,]*([0-9]{4})[\.,\s_]*([0-9]{4})', text)
            if not sm_match:
                sm_match = re.search(r'[\$S§s]M[\s\.:,]*([0-9]{4,8})', text)
                val = sm_match.group(1) if sm_match else ""
                sm = f"SM{val[:4]}.{val[4:]}" if len(val) == 8 else (f"SM{val}" if val else None)
            else:
                sm = f"SM{sm_match.group(1)}.{sm_match.group(2)}"
                
            date_match = re.search(r'Ng[aàeè]y[\s;:,-]*([0-9]{1,2})[\/\.\-]([0-9]{1,2})[\/\.\-]([0-9]{4})', text, re.IGNORECASE)
            if not date_match:
                date_match = re.search(r'([0-9]{1,2})[\/\.\-]([0-9]{1,2})[\/\.\-]([0-9]{4})', text)
                
            if date_match:
                d_val = int(date_match.group(1))
                m_val = int(date_match.group(2))
                y_val = date_match.group(3)
                dt = f"{d_val:02d}/{m_val:02d}/{y_val}"
            else:
                dt = None
                
            if sm and dt:
                records.append({
                    'sm': sm,
                    'date': dt,
                    'page': page_num
                })
                
        if page_callback:
            page_callback(page_num, total_pages)
            
    return records

def create_excel_multi_sheet(file_results):
    """Tạo file Excel với mỗi file PDF là 1 sheet riêng"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    fill_head = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_head = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_alt = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    border = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
    )
    headers = ["STT", "SỐ SM", "NGÀY", "SỐ TRANG CỦA SỐ SM"]
    
    existing_sheets = set()
    for file_name, records in file_results:
        sheet_title = sanitize_sheet_name(file_name, existing_sheets)
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        
        ws.row_dimensions[1].height = 26
        for idx, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=idx, value=h)
            c.fill = fill_head
            c.font = font_head
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
            
        for r_idx, r in enumerate(records, start=2):
            stt = r_idx - 1
            ws.row_dimensions[r_idx].height = 20
            row_vals = [stt, r['sm'], r['date'], r['page']]
            for c_idx, v in enumerate(row_vals, 1):
                c = ws.cell(row=r_idx, column=c_idx, value=v)
                c.font = Font(name="Calibri", size=11)
                c.border = border
                c.alignment = Alignment(horizontal="center", vertical="center")
                if r_idx % 2 == 1:
                    c.fill = fill_alt
                    
        ws.freeze_panes = "A2"
        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 25
        
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def render_auto_downloader(file_bytes, filename):
    """
    Tự động kích hoạt tải file về máy người dùng bằng JavaScript.
    Chạy trực tiếp từ cửa sổ trình duyệt chính.
    """
    b64 = base64.b64encode(file_bytes).decode()
    js_code = f"""
    <script>
        function triggerDownload() {{
            try {{
                var targetDoc = (window.parent && window.parent.document) ? window.parent.document : document;
                var a = targetDoc.createElement('a');
                a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}';
                a.download = '{filename}';
                targetDoc.body.appendChild(a);
                a.click();
                setTimeout(function() {{
                    try {{ targetDoc.body.removeChild(a); }} catch(err) {{}}
                }}, 1000);
            }} catch(e) {{
                var a2 = document.createElement('a');
                a2.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}';
                a2.download = '{filename}';
                document.body.appendChild(a2);
                a2.click();
            }}
        }}
        // Kích hoạt ngay lập tức
        setTimeout(triggerDownload, 300);
    </script>
    """
    components.html(js_code, height=0, width=0)

# ==================== GIAO DIỆN STREAMLIT TỐI GIẢN (1 TRANG) ====================
st.set_page_config(
    page_title="Trích Xuất Số SM - Tiền Phong", 
    page_icon="⚡", 
    layout="centered"
)

# Custom CSS: Vừa vặn 1 trang, không cần cuộn chuột
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0.5rem !important;
        max-width: 820px !important;
    }
    
    .hero-title {
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
        font-size: 24px;
        font-weight: 800;
        background: linear-gradient(135deg, #1E3A8A 0%, #0284C7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-top: 0px;
        margin-bottom: 4px;
        text-align: center;
    }
    
    .hero-sub {
        font-size: 13px;
        color: #64748B;
        text-align: center;
        margin-bottom: 12px;
    }
    
    .metric-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 8px 12px;
        text-align: center;
        margin-bottom: 8px;
    }
    .metric-val {
        font-size: 17px;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-label {
        font-size: 11px;
        color: #64748B;
        font-weight: 500;
    }
    
    .success-banner {
        background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
        border: 1px solid #A7F3D0;
        border-radius: 12px;
        padding: 14px 18px;
        color: #065F46;
        margin-top: 10px;
        margin-bottom: 14px;
    }
</style>
""", unsafe_allow_html=True)

# Header gọn gàng
st.markdown("""
<div style="text-align: center;">
    <h1 class="hero-title">Trích Xuất Số SM & Ngày Phiếu Giao Hàng</h1>
    <p class="hero-sub">Mẫu Nhựa Tiền Phong • Tự động đếm trang • Xuất Excel nhiều sheet</p>
</div>
""", unsafe_allow_html=True)

# Khung Upload
uploaded_files = st.file_uploader(
    "Chọn hoặc kéo thả các file PDF scan vào đây:", 
    type=["pdf"], 
    accept_multiple_files=True,
    help="Có thể chọn nhiều file cùng lúc. Bấm dấu ✖ cạnh file để xóa nếu chọn nhầm."
)

if uploaded_files:
    total_files = len(uploaded_files)
    st.caption(f"📁 Đã chọn **{total_files}** file PDF.")
    
    # 1. NÚT "BẮT ĐẦU TRÍCH XUẤT" (Chỉ hiện khi CHƯA xử lý xong)
    if 'completed' not in st.session_state:
        if st.button("⚡ BẮT ĐẦU TRÍCH XUẤT", type="primary", use_container_width=True):
            
            with st.spinner("Đang chuẩn bị quét dữ liệu..."):
                pages_per_file = []
                for f in uploaded_files:
                    try:
                        pdf_temp = pdfium.PdfDocument(f.getvalue())
                        pages_per_file.append(len(pdf_temp))
                    except Exception:
                        pages_per_file.append(1)
                total_pages_all = sum(pages_per_file)

            progress_bar = st.progress(0.0)
            
            # Cột 1: File (Đầu) | Cột 2: Tiến độ | Cột 3: Thời gian (Cuối)
            col_file, col_progress, col_time = st.columns(3)
            file_placeholder = col_file.empty()
            progress_placeholder = col_progress.empty()
            time_placeholder = col_time.empty()
            
            start_time = time.time()
            file_results = []
            tracker = {'pages_done': 0}
            total_extracted_all = 0
            
            for f_idx, up_file in enumerate(uploaded_files):
                file_name = up_file.name
                file_bytes = up_file.read()
                f_pages = pages_per_file[f_idx]
                
                def on_page_done(page_num, total_p):
                    tracker['pages_done'] += 1
                    pages_done = tracker['pages_done']
                    
                    elapsed = time.time() - start_time
                    avg_time = elapsed / pages_done
                    remaining_pages = max(total_pages_all - pages_done, 0)
                    remaining_sec = max(int(remaining_pages * avg_time), 1) if pages_done < total_pages_all else 0
                    
                    pct = min(pages_done / total_pages_all, 1.0)
                    progress_bar.progress(pct)
                    
                    # Cột 1: THÔNG TIN FILE (Ở ĐẦU)
                    file_placeholder.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-val" style="color: #D97706;">File {f_idx + 1}/{total_files}</div>
                        <div class="metric-label">{file_name[:16]}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Cột 2: TIẾN ĐỘ (%)
                    progress_placeholder.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-val" style="color: #059669;">{int(pct * 100)}%</div>
                        <div class="metric-label">Trang {pages_done}/{total_pages_all}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Cột 3: THỜI GIAN CÒN LẠI (Ở CUỐI - CẢ PHÚT VÀ GIÂY)
                    time_placeholder.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-val" style="color: #2563EB;">⏳ {format_time_str(remaining_sec)}</div>
                        <div class="metric-label">Thời gian còn lại</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                records = extract_from_single_pdf(file_bytes, page_callback=on_page_done)
                file_results.append((file_name, records))
                total_extracted_all += len(records)
                
            total_time_spent = int(time.time() - start_time)
            progress_bar.progress(1.0)
            
            # Tạo file Excel - TÊN FILE ĐỒNG NHẤT: PDF-TO-EXCEL.xlsx
            excel_bytes = create_excel_multi_sheet(file_results)
            out_name = "PDF-TO-EXCEL.xlsx"
            
            # Lưu session để chuyển trạng thái và hiển thị màn hình hoàn thành
            st.session_state['excel_data'] = excel_bytes
            st.session_state['excel_name'] = out_name
            st.session_state['total_files'] = total_files
            st.session_state['total_extracted'] = total_extracted_all
            st.session_state['total_time_str'] = format_time_str(total_time_spent)
            st.session_state['completed'] = True
            st.rerun()

# ==================== MÀN HÌNH HOÀN THÀNH ====================
if st.session_state.get('completed', False):
    out_name = st.session_state['excel_name']
    excel_bytes = st.session_state['excel_data']
    
    # Kích hoạt tự động tải file về máy
    render_auto_downloader(excel_bytes, out_name)
    
    st.markdown(f"""
    <div class="success-banner">
        <h4 style="margin: 0 0 6px 0; color: #065F46;">🎉 XỬ LÝ HOÀN TẤT ({st.session_state['total_time_str']})</h4>
        <div style="font-size: 13px; line-height: 1.5;">
            • Đã quét: <b>{st.session_state['total_files']}</b> file PDF &nbsp;|&nbsp; Tìm thấy: <b>{st.session_state['total_extracted']}</b> phiếu SM<br>
            • 📥 <b>File Excel <code>{out_name}</code> đang được tự động tải về máy của bạn!</b>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Nút XỬ LÝ FILE MỚI (Refresh trang về ban đầu)
    if st.button("🔄 XỬ LÝ FILE MỚI (LÀM MỚI TRANG)", type="primary", use_container_width=True):
        for k in ['excel_data', 'excel_name', 'total_files', 'total_extracted', 'total_time_str', 'completed']:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    # Link tải dự phòng nhỏ bên dưới phòng khi trình duyệt chặn tự động tải
    st.markdown("<div style='text-align: center; margin-top: 6px;'>", unsafe_allow_html=True)
    st.download_button(
        label="👉 Nếu trình duyệt chặn tự động tải, bấm vào đây để tải file",
        data=excel_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("</div>", unsafe_allow_html=True)
