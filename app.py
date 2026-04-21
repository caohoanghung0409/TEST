import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
import pytesseract
from PIL import Image
import numpy as np

# Cấu hình giao diện
st.set_page_config(page_title="Trích xuất PDF Tiền Phong - OCR Mode", layout="wide")

st.title("🚀 Hệ thống trích xuất Nhựa Tiền Phong (Chế độ OCR)")
st.info("Chế độ này sẽ quét hình ảnh để đọc chữ, xử lý được cả file scan bị lỗi font hoặc bị ngược.")

uploaded_file = st.file_uploader("Upload file PDF scan của bạn", type="pdf")

def process_ocr(page):
    """Chuyển trang PDF thành ảnh và dùng OCR để đọc chữ"""
    # Chuyển trang PDF thành ảnh (độ phân giải 300dpi để rõ nét)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # Dùng Tesseract để đọc chữ (hỗ trợ tiếng Việt)
    # Nếu chạy trên Streamlit Cloud, mặc định đã có sẵn tesseract
    text = pytesseract.image_to_string(img, lang='vie+eng')
    return text.upper()

if uploaded_file:
    with st.spinner('Đang dùng OCR quét hình ảnh (quá trình này có thể mất vài giây mỗi trang)...'):
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        results = []
        stt_counter = 1
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            
            # Quét lần 1 (Góc 0 độ)
            text_upper = process_ocr(page)
            
            # Kiểm tra tiêu đề
            keywords = ["NHỰA THIẾU NIÊN TIỀN PHONG", "ĐỒNG AN 2", "XÔ VIẾT NGHỆ TĨNH", "HÒA PHÚ"]
            is_valid = any(kw in text_upper for kw in keywords)
            
            # Nếu không thấy, xoay 180 độ và quét lại lần 2
            if not is_valid:
                page.set_rotation(180)
                text_upper = process_ocr(page)
                is_valid = any(kw in text_upper for kw in keywords)
            
            if is_valid:
                # 1. Trích xuất Ngày
                date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text_upper)
                ngay = date_match.group(1) if date_match else ""

                # 2. Trích xuất PR/SO
                pr_match = re.search(r"PR[\s]?(\d{4}[\d\.]*)", text_upper)
                so_match = re.search(r"SO[\s]?(\d{4}[\d\.]*)", text_upper)
                
                pr_val = pr_match.group(0).replace(" ", "") if pr_match else ""
                so_val = so_match.group(0).replace(" ", "") if so_match else ""
                
                pr_so_final = f"{pr_val}/{so_val}" if (pr_val and so_val) else (pr_val or so_val)

                # 3. Trích xuất SM
                sm_match = re.search(r"SM[\s]?([\d\.,]+)", text_upper)
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

        doc.close()

        if results:
            df = pd.DataFrame(results)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            st.success(f"Thành công! Đã tìm thấy {len(df)} phiếu.")
            st.table(df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Result')
            
            st.download_button(
                label="📥 Tải file Excel kết quả",
                data=output.getvalue(),
                file_name=f"KetQua_OCR_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Không tìm thấy dữ liệu trên ảnh. Vui lòng kiểm tra độ nét của file scan.")
