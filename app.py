import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Trích xuất PDF Tiền Phong", layout="wide")

st.title("🚀 Công cụ trích xuất dữ liệu Nhựa Tiền Phong")
st.write("Cơ chế: Vét cạn ký tự (Chống lỗi font và khoảng trắng do scan)")

uploaded_file = st.file_uploader("Upload file PDF", type="pdf")

def extract_advanced(text_raw):
    # Bước 1: Xóa sạch mọi loại khoảng trắng, xuống dòng để dồn chữ thành 1 khối duy nhất
    # Điều này giúp trị dứt điểm lỗi chữ bị rời rạc kiểu "P R 2 6 0 4"
    text_compact = re.sub(r'\s+', '', text_raw).upper()
    
    # Bước 2: Kiểm tra tiêu đề (Tìm từ khóa trong khối chữ đã dồn)
    keywords = ["NHỰATHIẾUNIÊNTIỀNPHONG", "NHUATHIEUNIENTIENPHONG", "ĐỒNGAN2", "DONGAN2"]
    if any(kw in text_compact for kw in keywords):
        
        # 1. Trích xuất Ngày (Tìm trong text_raw để lấy định dạng dd/mm/yyyy chuẩn)
        date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text_raw)
        ngay = date_match.group(1) if date_match else ""

        # 2. Trích xuất PR và SO từ khối chữ dồn
        # Tìm PR theo sau là 4 số năm (2604) và các số/dấu chấm tiếp theo
        pr_match = re.search(r"PR(260[3-4][\d\.]*)", text_compact)
        so_match = re.search(r"SO(260[3-4][\d\.]*)", text_compact)
        
        pr_val = f"PR{pr_match.group(1)}" if pr_match else ""
        so_val = f"SO{so_match.group(1)}" if so_match else ""
        
        pr_so_final = ""
        if pr_val and so_val:
            pr_so_final = f"{pr_val}/{so_val}"
        else:
            pr_so_final = pr_val if pr_val else so_val

        # 3. Trích xuất SM từ khối chữ dồn
        sm_match = re.search(r"SM(260[3-4][\d\.,]*)", text_compact)
        sm_val = f"SM{sm_match.group(1).replace(',', '.')}" if sm_match else ""

        if sm_val or pr_so_final:
            return {"SM": sm_val, "PR/SO": pr_so_final, "NGÀY": ngay}
            
    return None

if uploaded_file:
    with st.spinner('Hệ thống đang quét lớp chữ ẩn...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        data = []
        stt = 1
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Thử đọc 0 độ
            res = extract_advanced(page.get_text())
            
            # Nếu không thấy, thử xoay 180 độ
            if not res:
                page.set_rotation(180)
                res = extract_advanced(page.get_text())
            
            if res:
                res["STT"] = stt
                res["SỐ TRANG"] = i + 1
                data.append(res)
                stt += 1
        
        doc.close()

        if data:
            df = pd.DataFrame(data)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            st.success(f"Tìm thấy {len(df)} phiếu!")
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
            st.error("Hệ thống vẫn không đọc được lớp chữ ẩn của file này.")
            st.info("Vì file của bạn bôi đen được nhưng máy không đọc ra chữ đúng, hãy dùng cách cuối cùng: Mở PDF bằng Chrome -> Bấm Ctrl+P -> Chọn 'Save as PDF' rồi upload file mới đó lên đây.")
