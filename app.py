import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Trích xuất PDF Tiền Phong", page_icon="🚚", layout="wide")

st.title("📊 Hệ thống trích xuất dữ liệu Nhựa Tiền Phong (Bản Cao Cấp)")
st.info("Hỗ trợ đọc trang xuôi (0°) và trang ngược (180°). Tự động bỏ qua các trang không đúng tiêu đề.")

uploaded_file = st.file_uploader("Upload file PDF scan của bạn", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý dữ liệu...'):
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        results = []
        stt_counter = 1
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # --- BƯỚC 1: ĐỌC THỬ Ở GÓC 0 ĐỘ ---
            text = page.get_text()
            text_upper = text.upper()
            
            # Kiểm tra tiêu đề ở góc 0 độ
            is_tien_phong = "NHỰA THIẾU NIÊN TIỀN PHONG" in text_upper
            
            # --- BƯỚC 2: NẾU KHÔNG THẤY, THỬ XOAY 180 ĐỘ (Giả lập đọc ngược) ---
            if not is_tien_phong:
                # Xoay trang 180 độ
                page.set_rotation(180)
                text = page.get_text()
                text_upper = text.upper()
                is_tien_phong = "NHỰA THIẾU NIÊN TIỀN PHONG" in text_upper
            
            # --- BƯỚC 3: XỬ LÝ NẾU ĐÚNG TIÊU ĐỀ ---
            if is_tien_phong:
                # Làm sạch văn bản để quét Regex chính xác hơn
                text_clean = " ".join(text.split())
                
                # 1. Trích xuất NGÀY (dd/mm/yyyy)
                date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text_clean)
                ngay = date_match.group(1) if date_match else ""

                # 2. Trích xuất PR/SO
                # Tìm PR... và SO... linh hoạt
                pr_match = re.search(r"PR[\s]?(\d{4}[\d\.]*)", text_upper)
                so_match = re.search(r"SO[\s]?(\d{4}[\d\.]*)", text_upper)
                
                pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
                so_val = f"SO{so_match.group(1)}" if so_match else ""
                
                if pr_val and so_val:
                    pr_so_final = f"{pr_val}/{so_val}"
                else:
                    pr_so_final = pr_val if pr_val else so_val

                # 3. Trích xuất SM
                sm_match = re.search(r"SM[\s]?([\d\.,]+)", text_upper)
                sm_val = f"SM{sm_match.group(1).strip()}" if sm_match else ""

                # Chỉ lưu nếu tìm thấy dữ liệu định danh
                if sm_val or pr_so_final:
                    results.append({
                        "STT": stt_counter,
                        "SM": sm_val,
                        "PR/SO": pr_so_final,
                        "NGÀY": ngay,
                        "SỐ TRANG": i + 1
                    })
                    stt_counter += 1
            # Nếu không phải tiêu đề Nhựa Tiền Phong ở cả 2 chiều, bỏ qua luôn (Xử lý cực nhanh)

        doc.close()

        if results:
            df = pd.DataFrame(results)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            st.success(f"Đã xử lý xong! Tìm thấy {len(df)} phiếu hợp lệ.")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='KetQua')
            
            st.download_button(
                label="📥 Tải file Excel kết quả",
                data=output.getvalue(),
                file_name=f"KetQua_TienPhong_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Không tìm thấy dữ liệu phù hợp (Kể cả khi đã thử xoay trang).")
            st.warning("Gợi ý: Nếu file bôi đen được mà vẫn lỗi, có thể lớp chữ ẩn bị lỗi font cực nặng. Hãy thử dùng tính năng 'Save as PDF' của Chrome như mẹo trước đó.")
