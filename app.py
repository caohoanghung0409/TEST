import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

# Cấu hình giao diện
st.set_page_config(page_title="Trích xuất PDF Tiền Phong", page_icon="🚚", layout="wide")

st.title("📊 Hệ thống trích xuất dữ liệu Nhựa Tiền Phong")
st.write("Dữ liệu trích xuất: STT, SM, PR/SO, Ngày, Số trang")

uploaded_file = st.file_uploader("Upload file PDF scan của bạn", type="pdf")

if uploaded_file:
    with st.spinner('Đang phân tích dữ liệu...'):
        # Đọc dữ liệu từ file upload
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        results = []
        stt_counter = 1
        
        # Duyệt từng trang trong file PDF
        for i in range(len(doc)):
            page = doc.load_page(i)
            text = page.get_text()
            text_upper = text.upper()
            
            # ĐIỀU KIỆN LẤY TRANG: Nới lỏng theo các thông tin bạn cung cấp
            # Chỉ cần khớp tên công ty hoặc các thông tin địa chỉ đặc trưng
            check_header = any(key in text_upper for key in [
                "NHỰA THIẾU NIÊN TIỀN PHONG PHÍA NAM",
                "LÔ C2, KCN ĐỒNG AN 2",
                "HÒA PHÚ, TP.TDM",
                "135 XÔ VIẾT NGHỆ TĨNH"
            ])
            
            if check_header:
                # 1. Trích xuất NGÀY (Tìm định dạng dd/mm/yyyy)
                date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
                ngay = date_match.group(1) if date_match else ""

                # 2. Trích xuất PR/SO
                # Tìm các mã bắt đầu bằng PR hoặc SO, chấp nhận dấu gạch chéo ở giữa
                pr_so_pattern = r"(PR\d{4}[\d\.]*/SO\d{4}[\d\.]*|PR\d{4}[\d\.]*|SO\d{4}[\d\.]*)"
                pr_so_matches = re.findall(pr_so_pattern, text_upper)
                
                # Ưu tiên lấy mã có độ dài lớn nhất (thường là mã đầy đủ PR.../SO...)
                pr_so_val = max(pr_so_matches, key=len) if pr_so_matches else ""

                # 3. Trích xuất SM
                # Tìm chữ SM kèm các chữ số và dấu chấm/phẩy (ví dụ SM2604.3704)
                sm_match = re.search(r"SM\s?(\d{4}[\d\.,]*)", text_upper)
                sm_val = f"SM{sm_match.group(1).strip()}" if sm_match else ""

                # Chỉ lưu nếu tìm thấy ít nhất 1 loại mã số
                if sm_val or pr_so_val:
                    results.append({
                        "STT": stt_counter,
                        "SM": sm_val,
                        "PR/SO": pr_so_val,
                        "NGÀY": ngay,
                        "SỐ TRANG": i + 1
                    })
                    stt_counter += 1

        doc.close()

        if results:
            df = pd.DataFrame(results)
            # Đảm bảo thứ tự 5 cột như yêu cầu
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            
            st.success(f"Đã tìm thấy {len(df)} dòng dữ liệu hợp lệ!")
            st.dataframe(df, use_container_width=True)

            # Xuất file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='KetQua')
            
            st.download_button(
                label="📥 Tải file Excel ngay",
                data=output.getvalue(),
                file_name=f"Ket_qua_Tien_Phong_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Không tìm thấy dữ liệu. Hãy đảm bảo file PDF có lớp văn bản (bôi đen được).")

st.markdown("---")
st.caption("Công cụ tối ưu cho mẫu phiếu Nhựa Tiền Phong Phía Nam.")
