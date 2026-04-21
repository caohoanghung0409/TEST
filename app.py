import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Trích xuất PDF Tiền Phong", layout="wide")

st.title("⚡ Xử lý nhanh Phiếu Nhựa Tiền Phong")
st.write("Phiên bản tối ưu tốc độ, không dùng OCR để tránh treo máy.")

uploaded_file = st.file_uploader("Chọn file PDF", type="pdf")

def get_clean_data(text_raw):
    # Chuyển về chữ hoa và xóa các khoảng trắng thừa/xuống dòng dính liền
    text = " ".join(text_raw.split()).upper()
    
    # ĐIỀU KIỆN LẤY TRANG: Chỉ cần chứa tên công ty hoặc địa chỉ đặc trưng
    if "NHỰA THIẾU NIÊN TIỀN PHONG" in text or "ĐỒNG AN 2" in text or "XÔ VIẾT NGHỆ TĨNH" in text:
        
        # 1. Trích xuất Ngày (dd/mm/yyyy)
        date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
        ngay = date_match.group(1) if date_match else ""

        # 2. Trích xuất PR và SO (Tìm độc lập rồi ghép lại)
        pr_match = re.search(r"PR\s?(\d{4}[\d\.]*)", text)
        so_match = re.search(r"SO\s?(\d{4}[\d\.]*)", text)
        
        pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
        so_val = f"SO{so_match.group(1)}" if so_match else ""
        
        pr_so_final = ""
        if pr_val and so_val:
            pr_so_final = f"{pr_val}/{so_val}"
        else:
            pr_so_final = pr_val if pr_val else so_val

        # 3. Trích xuất SM (Lấy mã SM kèm số phía sau)
        # Regex này bắt được cả SM2604.0416 hoặc SM2604,0416
        sm_match = re.search(r"SM\s?(\d{4}[\d\.,]*)", text)
        sm_val = f"SM{sm_match.group(1).replace(',', '.')}" if sm_match else ""

        if sm_val or pr_so_final:
            return {"SM": sm_val, "PR/SO": pr_so_final, "NGÀY": ngay}
    return None

if uploaded_file:
    with st.spinner('Đang xử lý siêu tốc...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        final_list = []
        stt = 1
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Thử đọc 0 độ
            res = get_clean_data(page.get_text())
            
            # Nếu không thấy, thử xoay 180 độ
            if not res:
                page.set_rotation(180)
                res = get_clean_data(page.get_text())
                
            if res:
                res["STT"] = stt
                res["SỐ TRANG"] = i + 1
                final_list.append(res)
                stt += 1
        
        doc.close()

        if final_list:
            df = pd.DataFrame(final_list)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]] # Sắp xếp đúng 5 cột
            
            st.success(f"Xử lý xong! Tìm thấy {len(df)} phiếu.")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
            
            st.download_button(
                label="📥 Tải file Excel",
                data=output.getvalue(),
                file_name=f"KetQua_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Không tìm thấy trang nào đúng tiêu đề 'Nhựa Tiền Phong'.")
            st.info("Lưu ý: Nếu bạn bôi đen được chữ nhưng không ra kết quả, có thể lớp chữ ẩn bị lỗi. Bạn hãy dùng Chrome để 'In -> Lưu PDF' rồi upload lại nhé.")
