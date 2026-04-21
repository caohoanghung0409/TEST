import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

# Cấu hình giao diện
st.set_page_config(page_title="Lọc Phiếu Nhựa Tiền Phong", page_icon="🚚", layout="wide")

# CSS để giao diện chuyên nghiệp hơn
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; background-color: #007bff; color: white; }
    </style>
    """, unsafe_access_allowed=True)

st.title("🚚 Công cụ Trích xuất Phiếu Giao Hàng Nhựa Tiền Phong")
st.write("Dành riêng cho mẫu phiếu: CÔNG TY CỔ PHẦN NHỰA THIẾU NIÊN TIỀN PHONG PHÍA NAM")

uploaded_file = st.file_uploader("Kéo thả file PDF vào đây", type="pdf")

if uploaded_file:
    with st.spinner('Hệ thống đang quét dữ liệu...'):
        # Đọc PDF từ bộ nhớ
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        results = []
        stt = 1
        
        # Duyệt từng trang
        for i in range(len(doc)):
            page = doc.load_page(i)
            text = page.get_text()
            text_upper = text.upper()
            
            # ĐIỀU KIỆN 1: Kiểm tra đúng tên công ty và địa chỉ Tiền Phong Phía Nam
            header_keywords = [
                "NHỰA THIẾU NIÊN TIỀN PHONG PHÍA NAM",
                "LÔ C2, KCN ĐỒNG AN 2"
            ]
            
            if any(key in text_upper for key in header_keywords):
                
                # ĐIỀU KIỆN 2: Trích xuất Ngày (dd/mm/yyyy)
                date_match = re.search(r"(\d{2}/\d{2}/\d{4})", text)
                ngay = date_match.group(1) if date_match else ""

                # ĐIỀU KIỆN 3: Trích xuất PR/SO
                # Tìm các định dạng như PR2604.0416/SO2604.0853 hoặc PR2604.0416
                pr_so_match = re.search(r"(PR\d{4}[^\s]*SO\d{4}[^\s]*|PR\d{4}[^\s]*|SO\d{4}[^\s]*)", text)
                pr_so_val = pr_so_match.group(1) if pr_so_match else ""

                # ĐIỀU KIỆN 4: Trích xuất SM
                # Tìm mã bắt đầu bằng SM kèm chuỗi số phía sau
                sm_match = re.search(r"(SM\d{4}[\.,]\d+)", text)
                sm_val = sm_match.group(1) if sm_match else ""

                # Chỉ lấy dữ liệu nếu tìm thấy ít nhất mã SM hoặc PR/SO
                if sm_val or pr_so_val:
                    results.append({
                        "STT": stt,
                        "SM": sm_val,
                        "PR/SO": pr_so_val,
                        "NGÀY": ngay,
                        "SỐ TRANG": i + 1
                    })
                    stt += 1

        doc.close()

        if results:
            df = pd.DataFrame(results)
            
            # Sắp xếp lại thứ tự cột cho đúng yêu cầu
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            
            st.success(f"Phân tích hoàn tất! Tìm thấy {len(df)} phiếu giao hàng.")
            
            # Hiển thị bảng dữ liệu
            st.dataframe(df, use_container_width=True)

            # Tạo file Excel trong bộ nhớ
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
            
            # Nút tải file
            st.download_button(
                label="📥 Tải File Excel (5 Cột)",
                data=output.getvalue(),
                file_name=f"KetQua_TienPhong_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Không tìm thấy dữ liệu phù hợp trong file PDF này. Vui lòng kiểm tra lại file scan.")

st.markdown("---")
st.caption("Công cụ hỗ trợ xử lý dữ liệu cá nhân - Tốc độ cao - Bảo mật")
