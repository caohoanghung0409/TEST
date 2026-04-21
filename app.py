import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Fix Trích Xuất Tiền Phong", layout="wide")

st.title("✅ Hệ thống trích xuất Nhựa Tiền Phong (Bản Fix Dứt Điểm)")
st.write("Cơ chế: Sắp xếp lại văn bản theo tọa độ để trị lỗi chữ bị nhảy dòng/dính chữ.")

uploaded_file = st.file_uploader("Upload file PDF", type="pdf")

def solve_extraction(page, page_num):
    # Lấy văn bản theo dạng "dict" để có tọa độ từng chữ
    blocks = page.get_text("dict")["blocks"]
    full_text = ""
    
    # Gom tất cả các đoạn chữ lại theo thứ tự đọc tự nhiên
    for b in blocks:
        if "lines" in b:
            for l in b["lines"]:
                for s in l["spans"]:
                    full_text += s["text"] + " "
    
    text_upper = " ".join(full_text.split()).upper()
    
    # Kiểm tra tiêu đề (Nới lỏng tối đa: chỉ cần thấy chữ TIEN PHONG hoặc TIỀN PHONG)
    if "TIỀN PHONG" in text_upper or "TIEN PHONG" in text_upper or "ĐỒNG AN 2" in text_upper:
        
        # 1. Trích xuất Ngày
        date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text_upper)
        ngay = date_match.group(1) if date_match else ""

        # 2. Trích xuất PR và SO (Tìm độc lập)
        # Bắt các chuỗi bắt đầu bằng PR/SO và có dãy số 26...
        pr_match = re.search(r"PR\s?(26\d{2}[\d\.]*)", text_upper)
        so_match = re.search(r"SO\s?(26\d{2}[\d\.]*)", text_upper)
        
        pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
        so_val = f"SO{so_match.group(1)}" if so_match else ""
        
        pr_so_final = ""
        if pr_val and so_val:
            pr_so_final = f"{pr_val}/{so_val}"
        else:
            pr_so_final = pr_val if pr_val else so_val

        # 3. Trích xuất SM
        sm_match = re.search(r"SM\s?(26\d{2}[\d\.,]*)", text_upper)
        sm_val = f"SM{sm_match.group(1).replace(',', '.')}" if sm_match else ""

        if sm_val or pr_so_final:
            return {"STT": 0, "SM": sm_val, "PR/SO": pr_so_final, "NGÀY": ngay, "SỐ TRANG": page_num}
    return None

if uploaded_file:
    with st.spinner('Đang nhặt ký tự và ghép mã...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        data = []
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            # Thử hướng 0 độ
            res = solve_extraction(page, i + 1)
            
            # Nếu không thấy, xoay 180 độ thử lại
            if not res:
                page.set_rotation(180)
                res = solve_extraction(page, i + 1)
                
            if res:
                data.append(res)
        
        doc.close()

        if data:
            df = pd.DataFrame(data)
            # Cập nhật lại STT
            df['STT'] = range(1, len(df) + 1)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            
            st.success(f"Đã trích xuất thành công {len(df)} dòng dữ liệu!")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Result')
            
            st.download_button(
                label="📥 Tải File Excel 5 Cột",
                data=output.getvalue(),
                file_name=f"KetQua_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Lỗi: Hệ thống không tìm thấy tiêu đề 'Nhựa Tiền Phong'.")
