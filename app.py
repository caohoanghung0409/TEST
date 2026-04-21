import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Fix Triệt Để Tiền Phong", layout="wide")

st.title("✅ Trích xuất Nhựa Tiền Phong - Bản Fix Lỗi Nhận Diện")
st.info("Cơ chế: Quét toàn bộ văn bản (Global Scan) - Không quan trọng vị trí tiêu đề.")

uploaded_file = st.file_uploader("Upload file PDF của bạn", type="pdf")

def get_data_v2(text_content, page_num):
    # Chuẩn hóa văn bản: Xóa mọi dấu xuống dòng và khoảng trắng thừa thành 1 dòng duy nhất
    raw_text = " ".join(text_content.split())
    text_upper = raw_text.upper()
    
    # KIỂM TRA ĐIỀU KIỆN (Chỉ cần có 1 trong các từ khóa này ở bất cứ đâu)
    keywords = ["NHỰA THIẾU NIÊN TIỀN PHONG", "NHUATHIEUNIENTIENPHONG", "ĐỒNG AN 2", "TIEN PHONG PLASTIC"]
    if any(kw in text_upper for kw in keywords):
        
        # 1. Trích xuất Ngày (dd/mm/yyyy)
        date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", raw_text)
        ngay = date_match.group(1) if date_match else ""

        # 2. Trích xuất PR/SO
        # Tìm PR... và SO... linh hoạt nhất có thể
        pr_match = re.search(r"PR\s?(\d{4}[\d\.]*)", text_upper)
        so_match = re.search(r"SO\s?(\d{4}[\d\.]*)", text_upper)
        
        pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
        so_val = f"SO{so_match.group(1)}" if so_match else ""
        
        pr_so_final = ""
        if pr_val and so_val:
            pr_so_final = f"{pr_val}/{so_val}"
        else:
            pr_so_final = pr_val if pr_val else so_val

        # 3. Trích xuất SM
        sm_match = re.search(r"SM\s?(\d{4}[\d\.,]*)", text_upper)
        sm_val = ""
        if sm_match:
            sm_val = f"SM{sm_match.group(1).replace(',', '.')}"

        # Chỉ lấy nếu có mã SM hoặc PR/SO
        if sm_val or pr_so_final:
            return {
                "SM": sm_val,
                "PR/SO": pr_so_final,
                "NGÀY": ngay,
                "SỐ TRANG": page_num
            }
    return None

if uploaded_file:
    with st.spinner('Đang phân tích dữ liệu...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        results = []
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Thử hướng 0 độ
            content = page.get_text("text")
            res = get_data_v2(content, i + 1)
            
            # Nếu không thấy, xoay 180 độ
            if not res:
                page.set_rotation(180)
                content_rotated = page.get_text("text")
                res = get_data_v2(content_rotated, i + 1)
            
            if res:
                results.append(res)

        doc.close()

        if results:
            df = pd.DataFrame(results)
            # Thêm STT
            df.insert(0, 'STT', range(1, len(df) + 1))
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            
            st.success(f"Đã tìm thấy {len(df)} phiếu!")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            st.download_button(label="📥 Tải file Excel", data=output.getvalue(), 
                             file_name="KetQua_TienPhong.xlsx", 
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Vẫn không tìm thấy dữ liệu. Lớp chữ ẩn trong file scan của bạn bị lỗi vị trí quá nặng.")
            st.warning("Giải pháp cuối cùng: Bạn hãy dùng Chrome mở file -> In -> Lưu PDF (Save as PDF). File mới này sẽ có lớp chữ cực chuẩn cho máy đọc.")
