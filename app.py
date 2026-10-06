"""
ỨNG DỤNG STREAMLIT CHẠY TRÊN STREAMLIT COMMUNITY CLOUD (GITHUB)
Tương thích hoàn toàn với hệ thống Linux của Streamlit Cloud và Windows.
"""

import os
import sys
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pypdfium2 as pdfium
import pytesseract
from PIL import Image
import pandas as pd
import streamlit as st

# Tự động nhận diện Tesseract OCR trên Windows (nếu chạy local) hoặc Linux (trên Cloud)
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
    # Trên Streamlit Cloud (Linux), tesseract tự động nằm trong PATH của hệ thống

setup_tesseract()

def sanitize_sheet_name(name, existing_names):
    """Quy chuẩn tên Sheet trong Excel (tối đa 31 ký tự, không chứa ký tự cấm)"""
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

def extract_from_single_pdf(file_bytes, file_name, progress_callback=None):
    """Quét dữ liệu từng trang của 1 file PDF"""
    pdf = pdfium.PdfDocument(file_bytes)
    total_pages = len(pdf)
    records = []
    
    for page_idx in range(total_pages):
        page_num = page_idx + 1
        if progress_callback:
            progress_callback(page_num, total_pages)
            
        page = pdf[page_idx]
        bitmap = page.render(scale=2.0)
        img = bitmap.to_pil()
        w, h = img.size
        
        crop = img.crop((0, 0, w, int(h * 0.30)))
        try:
            text = pytesseract.image_to_string(crop, lang='eng')
        except Exception as e:
            st.error(f"Lỗi nhận diện OCR: {e}")
            return None
            
        text_u = text.upper()
        is_tp = ('TIEN PHONG' in text_u or 'THIEU NIEN' in text_u or 'NHUATIENPHONG' in text_u)
        if not is_tp:
            continue
            
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
            
    return records

def create_excel_multi_sheet(file_results):
    """Ghi tất cả kết quả vào 1 file Excel, mỗi file PDF là 1 sheet riêng"""
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

# ==================== GIAO DIỆN STREAMLIT ====================
st.set_page_config(
    page_title="Trích Xuất Số SM - Tiền Phong", 
    page_icon="📑", 
    layout="centered"
)

st.title("📑 Công Cụ Trích Xuất Số SM & Ngày (Nhựa Tiền Phong)")
st.caption("Công cụ trích xuất dữ liệu trực tuyến. Chọn 1 hoặc nhiều file PDF scan để xuất ra Excel.")

uploaded_files = st.file_uploader(
    "Chọn một hoặc nhiều file PDF scan cần xử lý:", 
    type=["pdf"], 
    accept_multiple_files=True,
    help="Có thể kéo thả nhiều file vào đây. Bấm dấu ✖ cạnh tên file để xóa nếu chọn nhầm."
)

if uploaded_files:
    st.info(f"📁 Đã chọn **{len(uploaded_files)}** file PDF.")
    
    if st.button("🚀 Bắt Đầu Trích Xuất", type="primary", use_container_width=True):
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        file_results = []
        total_files = len(uploaded_files)
        total_extracted_all = 0
        
        for f_idx, up_file in enumerate(uploaded_files):
            file_name = up_file.name
            file_bytes = up_file.read()
            
            def update_progress(page, total_p):
                overall = (f_idx / total_files) + (page / total_p) / total_files
                progress_bar.progress(min(overall, 1.0))
                status_text.text(f"Đang quét file [{f_idx + 1}/{total_files}]: {file_name} - Trang {page}/{total_p}...")
                
            records = extract_from_single_pdf(file_bytes, file_name, progress_callback=update_progress)
            if records is None:
                st.stop()
                
            file_results.append((file_name, records))
            total_extracted_all += len(records)
            
        progress_bar.progress(1.0)
        status_text.empty()
        
        excel_bytes = create_excel_multi_sheet(file_results)
        
        if total_files == 1:
            out_name = f"{os.path.splitext(uploaded_files[0].name)[0]}_KetQua_SM.xlsx"
        else:
            out_name = "TongHop_So_SM_TienPhong.xlsx"
            
        st.success(f"🎉 **Hoàn thành xử lý {total_files} file PDF!** Tổng cộng tìm thấy **{total_extracted_all}** phiếu SM hợp lệ.")
        
        st.download_button(
            label=f"📥 BẤM VÀO ĐÂY ĐỂ TẢI FILE EXCEL KẾT QUẢ ({out_name})",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
