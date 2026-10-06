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

def trigger_auto_download(file_bytes, filename):
    """Tự động kích hoạt tải file về máy qua JavaScript (vượt qua sandbox iframe)"""
    b64 = base64.b64encode(file_bytes).decode()
    js_code = f"""
    <script>
        (function() {{
            try {{
                var doc = window.parent.document || document;
                var a = doc.createElement('a');
                a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}';
                a.download = '{filename}';
                doc.body.appendChild(a);
                a.click();
                doc.body.removeChild(a);
            }} catch(e) {{
                var a = document.createElement('a');
                a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}';
                a.download = '{filename}';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }}
        }})();
    </script>
    """
    components.html(js_code, height=0, width=0)

# ==================== GIAO DIỆN HIỆN ĐẠI (STREAMLIT) ====================
st.set_page_config(
    page_title="Trích Xuất Số SM - Tiền Phong", 
    page_icon="⚡", 
    layout="centered"
)

# Custom CSS cho phong cách hiện đại (Modern Clean SaaS)
st.markdown("""
<style>
    /* Ẩn header và menu mặc định */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .main {
        background-color: #F8FAFC;
    }
    
    /* Khung card bo tròn thanh lịch */
    .app-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 24px 28px;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    
    /* Tiêu đề gradient */
    .hero-title {
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
        font-size: 28px;
        font-weight: 800;
        background: linear-gradient(135deg, #1E3A8A 0%, #0284C7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 6px;
        text-align: center;
    }
    
    .hero-sub {
        font-size: 14px;
        color: #64748B;
        text-align: center;
        margin-bottom: 24px;
    }
    
    /* Huy hiệu badge nhỏ */
    .badge {
        display: inline-block;
        background: #EFF6FF;
        color: #2563EB;
        font-size: 12px;
        font-weight: 600;
        padding: 4px 12px;
        border-radius: 9999px;
        border: 1px solid #DBEAFE;
        margin-bottom: 8px;
    }
    
    /* Metric hộp nhỏ */
    .metric-box {
        background: #F1F5F9;
        border-radius: 12px;
        padding: 12px 16px;
        text-align: center;
    }
    .metric-val {
        font-size: 22px;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-label {
        font-size: 12px;
        color: #64748B;
        font-weight: 500;
    }
    
    /* Thành công banner */
    .success-banner {
        background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
        border: 1px solid #A7F3D0;
        border-radius: 14px;
        padding: 18px 20px;
        color: #065F46;
        margin-top: 15px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div style="text-align: center;">
    <span class="badge">🚀 CÔNG CỤ TỰ ĐỘNG HÓA</span>
    <h1 class="hero-title">Trích Xuất Số SM & Ngày Phiếu Giao Hàng</h1>
    <p class="hero-sub">Nhận diện mẫu Nhựa Tiền Phong • Tự động đếm trang • Xuất Excel nhiều sheet</p>
</div>
""", unsafe_allow_html=True)

# Khung Upload
uploaded_files = st.file_uploader(
    "Chọn hoặc kéo thả các file PDF scan vào đây:", 
    type=["pdf"], 
    accept_multiple_files=True,
    help="Có thể chọn nhiều file cùng lúc. Bấm dấu ✖ bên cạnh tên file để xóa nếu chọn nhầm."
)

if uploaded_files:
    total_files = len(uploaded_files)
    st.caption(f"📁 Đã sẵn sàng **{total_files}** file PDF để xử lý.")
    
    # Nút bấm Bắt đầu trích xuất
    if st.button("⚡ BẮT ĐẦU TRÍCH XUẤT", type="primary", use_container_width=True):
        
        # 1. Đếm tổng số trang trước để tính toán thời gian chính xác
        with st.spinner("Đang chuẩn bị và tính toán khối lượng công việc..."):
            pages_per_file = []
            for f in uploaded_files:
                try:
                    pdf_temp = pdfium.PdfDocument(f.getvalue())
                    pages_per_file.append(len(pdf_temp))
                except Exception:
                    pages_per_file.append(1)
            total_pages_all = sum(pages_per_file)

        # 2. Hiển thị Dashboard thời gian & Tiến trình
        progress_bar = st.progress(0.0)
        
        # 3 cột thông số động
        col_time, col_progress, col_status = st.columns(3)
        time_placeholder = col_time.empty()
        progress_placeholder = col_progress.empty()
        status_placeholder = col_status.empty()
        
        # Ước tính ban đầu: trung bình 0.75s mỗi trang scan
        est_total_seconds = max(int(total_pages_all * 0.75), 3)
        start_time = time.time()
        
        file_results = []
        pages_processed_so_far = 0
        total_extracted_all = 0
        
        for f_idx, up_file in enumerate(uploaded_files):
            file_name = up_file.name
            file_bytes = up_file.read()
            f_pages = pages_per_file[f_idx]
            
            def on_page_done(page_num, total_p):
                nonlocal pages_processed_so_far
                pages_processed_so_far += 1
                
                elapsed = time.time() - start_time
                avg_time = elapsed / pages_processed_so_far
                remaining_pages = max(total_pages_all - pages_processed_so_far, 0)
                remaining_sec = max(int(remaining_pages * avg_time), 1) if pages_processed_so_far < total_pages_all else 0
                
                pct = min(pages_processed_so_far / total_pages_all, 1.0)
                progress_bar.progress(pct)
                
                # Cập nhật số giây đếm ngược từ từ về 1s
                time_placeholder.markdown(f"""
                <div class="metric-box">
                    <div class="metric-val" style="color: #2563EB;">⏳ {remaining_sec}s</div>
                    <div class="metric-label">Thời gian còn lại</div>
                </div>
                """, unsafe_allow_html=True)
                
                progress_placeholder.markdown(f"""
                <div class="metric-box">
                    <div class="metric-val" style="color: #059669;">{int(pct * 100)}%</div>
                    <div class="metric-label">Trang {pages_processed_so_far}/{total_pages_all}</div>
                </div>
                """, unsafe_allow_html=True)
                
                status_placeholder.markdown(f"""
                <div class="metric-box">
                    <div class="metric-val" style="color: #D97706;">File {f_idx + 1}/{total_files}</div>
                    <div class="metric-label">Đang quét: {file_name[:12]}...</div>
                </div>
                """, unsafe_allow_html=True)
                
            records = extract_from_single_pdf(file_bytes, page_callback=on_page_done)
            file_results.append((file_name, records))
            total_extracted_all += len(records)
            
        total_time_spent = int(time.time() - start_time)
        progress_bar.progress(1.0)
        
        # 3. Tạo file Excel
        excel_bytes = create_excel_multi_sheet(file_results)
        
        if total_files == 1:
            out_name = f"{os.path.splitext(uploaded_files[0].name)[0]}_KetQua_SM.xlsx"
        else:
            out_name = "TongHop_So_SM_TienPhong.xlsx"
            
        # 4. Tự động kích hoạt Download về máy
        trigger_auto_download(excel_bytes, out_name)
        
        # 5. Lưu vào session_state để duy trì trạng thái hoàn thành
        st.session_state['excel_data'] = excel_bytes
        st.session_state['excel_name'] = out_name
        st.session_state['total_files'] = total_files
        st.session_state['total_extracted'] = total_extracted_all
        st.session_state['total_time'] = total_time_spent

# ==================== MÀN HÌNH HOÀN THÀNH ====================
if 'excel_data' in st.session_state:
    st.markdown(f"""
    <div class="success-banner">
        <h3 style="margin: 0 0 8px 0; color: #065F46;">🎉 XỬ LÝ HOÀN TẤT TRONG {st.session_state['total_time']} GIÂY!</h3>
        <p style="margin: 0; font-size: 14px;">
            • Đã xử lý: <b>{st.session_state['total_files']}</b> file PDF<br>
            • Tổng số phiếu SM hợp lệ: <b>{st.session_state['total_extracted']}</b> phiếu<br>
            • 📥 <b>File Excel <code>{st.session_state['excel_name']}</code> đang được tự động tải về máy của bạn!</b>
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Nút XỬ LÝ FILE MỚI (Refresh trang về ban đầu)
    if st.button("🔄 XỬ LÝ FILE MỚI (LÀM MỚI TRANG)", type="primary", use_container_width=True):
        for k in ['excel_data', 'excel_name', 'total_files', 'total_extracted', 'total_time']:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    # Link tải dự phòng nhỏ phòng trường hợp trình duyệt chặn tải tự động
    st.markdown("<div style='text-align: center; margin-top: 10px;'>", unsafe_allow_html=True)
    st.download_button(
        label="👉 Nếu trình duyệt chặn tải tự động, bấm vào đây để tải file",
        data=st.session_state['excel_data'],
        file_name=st.session_state['excel_name'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("</div>", unsafe_allow_html=True)
