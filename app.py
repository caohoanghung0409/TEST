import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Trích xuất PDF Tiền Phong", layout="wide")

st.title("📊 Hệ thống trích xuất dữ liệu Nhựa Tiền Phong")
st.markdown("---")

uploaded_file = st.file_uploader("Upload file PDF của bạn", type="pdf")

def extract_logic(text_block, page_num):
    # Chuẩn hóa văn bản: xóa khoảng trắng thừa, đưa về chữ hoa
    clean_text = " ".join(text_block.split()).upper()
    
    # 1. Kiểm tra tiêu đề (Chỉ cần chứa các chữ cái đặc trưng nhất)
    # Nếu không có chữ TIEN PHONG hoặc NHUA THIEU NIEN thì bỏ qua nhanh
    if not any(kw in clean_text for kw in ["TIỀN PHONG", "TIEN PHONG", "ĐỒNG AN 2", "XÔ VIẾT NGHỆ TĨNH"]):
        return None

    # 2. Tìm Ngày (dd/mm/yyyy)
    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", clean_text)
    ngay = date_match.group(1) if date_match else ""

    # 3. Tìm PR/SO
    # Tìm mã PR26... hoặc SO26...
    pr_match = re.search(r"PR\s?(\d{4}[\d\.]*)", clean_text)
    so_match = re.search(r"SO\s?(\d{4}[\d\.]*)", clean_text)
    
    pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
    so_val = f"SO{so_match.group(1)}" if so_match else ""
    
    pr_so_final = ""
    if pr_val and so_val:
        pr_so_final = f"{pr_val}/{so_val}"
    else:
        pr_so_final = pr_val if pr_val else so_val

    # 4. Tìm SM (Mã SM26...)
    sm_match = re.search(r"SM\s?(\d{4}[\d\.,]*)", clean_text)
    sm_val = f"SM{sm_match.group(1).replace(' ', '').replace(',', '.')}" if sm_match else ""

    if sm_val or pr_so_final:
        return {
            "SM": sm_val,
            "PR/SO": pr_so_final,
            "NGÀY": ngay,
            "SỐ TRANG": page_num
        }
    return None

if uploaded_file:
    with st.spinner('Đang quét dữ liệu...'):
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        final_results = []
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Thử đọc ở góc 0 độ
            page_text = page.get_text("text")
            res = extract_logic(page_text, i + 1)
            
            # Nếu không thấy, thử xoay 180 độ
            if not res:
                page.set_rotation(180)
                page_text_rotated = page.get_text("text")
                res = extract_logic(page_text_rotated, i + 1)
            
            if res:
                final_results.append(res)

        doc.close()

        if final_results:
            df = pd.DataFrame(final_results)
            # Thêm cột STT vào đầu
            df.insert(0, 'STT', range(1, len(df) + 1))
            
            st.success(f"Đã xử lý xong {len(df)} trang hợp lệ!")
            st.table(df)

            # Xuất Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
            
            st.download_button(
                label="📥 Tải file Excel kết quả",
                data=output.getvalue(),
                file_name=f"Ket_Qua_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Không tìm thấy dữ liệu. File này có lớp chữ ẩn bị hỏng hoàn toàn.")
            st.info("Vì bạn không muốn dùng OCR (tốn phí/chậm), hãy thử mở file bằng Chrome -> In -> Lưu PDF rồi upload lại nhé.")
