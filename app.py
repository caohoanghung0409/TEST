import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Fix Dứt Điểm Tiền Phong", layout="wide")

st.title("✅ Trích xuất Nhựa Tiền Phong - Bản Fix Dứt Điểm")
st.info("Cơ chế: Ghép mảnh ký tự (Char-Merging) - Chống lỗi scan chữ rời rạc.")

uploaded_file = st.file_uploader("Upload file PDF của bạn", type="pdf")

def solve_this_file(page, page_num):
    # Lấy text theo dạng từng từ đơn lẻ để tránh lỗi khoảng trắng ma
    words = page.get_text("words")
    # Ghép tất cả các từ lại, xóa sạch dấu xuống dòng
    full_text = " ".join([w[4] for w in words])
    text_upper = full_text.upper()
    
    # 1. KIỂM TRA TIÊU ĐỀ: Chỉ cần xuất hiện các từ khóa này ở BẤT KỲ ĐÂU trên trang
    keywords = ["NHỰA THIẾU NIÊN TIỀN PHONG", "TIEN PHONG", "ĐỒNG AN 2", "KHU PHỨC HỢP"]
    if any(kw in text_upper for kw in keywords):
        
        # 2. Trích xuất NGÀY (dd/mm/yyyy)
        date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", full_text)
        ngay = date_match.group(1) if date_match else ""

        # 3. Trích xuất PR/SO (Tìm cụm PR26... hoặc SO26...)
        # Mình dùng Regex nới lỏng để bắt được mã dù nó dính chữ hay cách quãng
        pr_match = re.search(r"PR\s?(26\d{2}[\d\.]*)", text_upper)
        so_match = re.search(r"SO\s?(26\d{2}[\d\.]*)", text_upper)
        
        pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
        so_val = f"SO{so_match.group(1)}" if so_match else ""
        
        # Kết hợp PR/SO
        if pr_val and so_val:
            pr_so_final = f"{pr_val}/{so_val}"
        else:
            pr_so_final = pr_val if pr_val else so_val

        # 4. Trích xuất SM (Mã SM26...)
        sm_match = re.search(r"SM\s?(26\d{2}[\d\.,]*)", text_upper)
        sm_val = f"SM{sm_match.group(1).replace(',', '.')}" if sm_match else ""

        # Nếu tìm thấy ít nhất 1 mã thì mới lấy dữ liệu
        if sm_val or pr_so_final:
            return {
                "SM": sm_val,
                "PR/SO": pr_so_final,
                "NGÀY": ngay,
                "SỐ TRANG": page_num
            }
    return None

if uploaded_file:
    with st.spinner('Đang nhặt ký tự và xử lý...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        results = []
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Thử hướng 0 độ
            res = solve_this_file(page, i + 1)
            
            # Nếu không thấy, thử xoay 180 độ
            if not res:
                page.set_rotation(180)
                res = solve_this_file(page, i + 1)
            
            if res:
                results.append(res)
        
        doc.close()

        if results:
            df = pd.DataFrame(results)
            # Thêm STT và sắp xếp cột
            df.insert(0, 'STT', range(1, len(df) + 1))
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            
            st.success(f"Đã trích xuất thành công {len(df)} dòng dữ liệu!")
            st.dataframe(df, use_container_width=True)

            # Xuất Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
            
            st.download_button(
                label="📥 Tải file Excel kết quả",
                data=output.getvalue(),
                file_name=f"KetQua_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Lỗi: Không tìm thấy dữ liệu Nhựa Tiền Phong. Có thể lớp chữ ẩn bị hỏng nặng.")
