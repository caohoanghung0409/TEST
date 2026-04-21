import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Trích xuất PDF Tiền Phong", page_icon="🚚", layout="wide")

st.title("📊 Hệ thống trích xuất dữ liệu Nhựa Tiền Phong")
st.info("Phiên bản tối ưu: Đọc dữ liệu từ file scan có lớp chữ ẩn bị lỗi.")

uploaded_file = st.file_uploader("Upload file PDF scan của bạn", type="pdf")

def clean_text(text):
    """Làm sạch văn bản, xử lý lỗi font và khoảng trắng"""
    text = re.sub(r'\s+', ' ', text)  # Gom nhiều khoảng trắng thành 1
    return text.strip()

if uploaded_file:
    with st.spinner('Đang phân tích dữ liệu...'):
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        results = []
        stt_counter = 1
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Thử đọc ở cả 2 hướng (0 độ và 180 độ)
            for rotation in [0, 180]:
                if rotation == 180:
                    page.set_rotation(180)
                
                text_raw = page.get_text("text")
                text_clean = clean_text(text_raw)
                text_upper = text_clean.upper()
                
                # Kiểm tra tiêu đề (Chỉ cần khớp 1 trong các từ khóa quan trọng)
                check_keywords = ["NHỰA THIẾU NIÊN TIỀN PHONG", "ĐỒNG AN 2", "XÔ VIẾT NGHỆ TĨNH", "HÒA PHÚ"]
                if any(kw in text_upper for kw in check_keywords):
                    
                    # 1. Lấy NGÀY
                    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text_clean)
                    ngay = date_match.group(1) if date_match else ""

                    # 2. Lấy PR/SO (Dùng Regex linh hoạt cho cả mã dính liền hoặc cách nhau)
                    # Tìm mã PR
                    pr_match = re.search(r"PR\s?(\d{4}[\d\.]*)", text_upper)
                    # Tìm mã SO
                    so_match = re.search(r"SO\s?(\d{4}[\d\.]*)", text_upper)
                    
                    pr_val = pr_match.group(0).replace(" ", "") if pr_match else ""
                    so_val = so_match.group(0).replace(" ", "") if so_match else ""
                    
                    pr_so_final = ""
                    if pr_val and so_val:
                        pr_so_final = f"{pr_val}/{so_val}"
                    else:
                        pr_so_final = pr_val if pr_val else so_val

                    # 3. Lấy SM
                    sm_match = re.search(r"SM\s?(\d{4}[\d\.,]*)", text_upper)
                    sm_val = sm_match.group(0).replace(" ", "").replace(",", ".") if sm_match else ""

                    if sm_val or pr_so_final:
                        results.append({
                            "STT": stt_counter,
                            "SM": sm_val,
                            "PR/SO": pr_so_final,
                            "NGÀY": ngay,
                            "SỐ TRANG": i + 1
                        })
                        stt_counter += 1
                        break # Đã tìm thấy dữ liệu ở hướng này, không cần xoay tiếp
        
        doc.close()

        if results:
            df = pd.DataFrame(results)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            # Loại bỏ các dòng trùng lặp nếu có do xoay trang
            df = df.drop_duplicates(subset=["SM", "PR/SO", "SỐ TRANG"])
            
            st.success(f"Đã tìm thấy {len(df)} phiếu giao hàng!")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            st.download_button(
                label="📥 Tải file Excel",
                data=output.getvalue(),
                file_name=f"KetQua_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Không tìm thấy dữ liệu. Lớp chữ ẩn trong file PDF này bị lỗi quá nặng.")
            st.info("Cách xử lý: Bạn hãy dùng điện thoại chụp lại hoặc scan lại file rõ hơn, hoặc dùng tính năng 'In -> Save as PDF' để tạo lại lớp chữ mới.")
