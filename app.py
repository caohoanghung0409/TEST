import streamlit as st
import numpy as np
from pdf2image import convert_from_bytes
import easyocr
import pandas as pd
import re
import io

st.set_page_config(page_title="Hệ thống Trích xuất Tiền Phong", layout="wide")
st.title("🚀 Bản Fix Dứt Điểm - Nhận diện hình ảnh (EasyOCR)")

uploaded_file = st.file_uploader("Upload file PDF scan của bạn", type="pdf")

# Khởi tạo bộ đọc OCR (chỉ chạy 1 lần để tiết kiệm thời gian)
@st.cache_resource
def load_reader():
    return easyocr.Reader(['vi', 'en'])

reader = load_reader()

def extract_data_from_image(img):
    # Chuyển ảnh sang dạng mảng để OCR
    results = reader.readtext(np.array(img), detail=0)
    full_text = " ".join(results).upper()
    
    # Logic tìm kiếm mã số (Pattern năm 260x)
    # 1. Tìm Ngày
    date_match = re.search(r"(\d{2}/\d{2}/\d{4})", full_text)
    ngay = date_match.group(1) if date_match else ""

    # 2. Tìm PR/SO
    pr = re.search(r"PR\s?(26\d{2}[\d\.]*)", full_text)
    so = re.search(r"SO\s?(26\d{2}[\d\.]*)", full_text)
    p_val = f"PR{pr.group(1)}" if pr else ""
    s_val = f"SO{so.group(1)}" if so else ""
    pr_so = f"{p_val}/{s_val}" if (p_val and s_val) else (p_val or s_val)

    # 3. Tìm SM
    sm = re.search(r"SM\s?(26\d{2}[\d\.,]*)", full_text)
    sm_val = f"SM{sm.group(1).replace(',', '.')}" if sm else ""

    return {"SM": sm_val, "PR/SO": pr_so, "NGÀY": ngay}

if uploaded_file:
    with st.spinner('Đang dùng mắt thần OCR quét hình ảnh...'):
        # Chuyển PDF thành ảnh (300 DPI để rõ nét nhất)
        images = convert_from_bytes(uploaded_file.read(), dpi=300)
        
        final_data = []
        for i, img in enumerate(images):
            # Quét dữ liệu trang xuôi
            data = extract_data_from_image(img)
            
            # Nếu không thấy mã, xoay 180 độ quét lại
            if not data["SM"] and not data["PR/SO"]:
                img_rotated = img.rotate(180)
                data = extract_data_from_image(img_rotated)
            
            if data["SM"] or data["PR/SO"]:
                data["STT"] = len(final_data) + 1
                data["SỐ TRANG"] = i + 1
                final_data.append(data)

        if final_data:
            df = pd.DataFrame(final_data)
            df = df[["STT", "SM", "PR/SO", "NGÀY", "SỐ TRANG"]]
            st.success(f"Đã xử lý xong! Tìm thấy {len(df)} phiếu.")
            st.table(df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Tải file Excel", output.getvalue(), "KetQua.xlsx")
        else:
            st.error("Vẫn không tìm thấy dữ liệu. Hãy kiểm tra độ nét của file scan.")
