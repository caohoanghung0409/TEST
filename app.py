import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Trích xuất PDF Tiền Phong", page_icon="🚚", layout="wide")

st.title("📊 Hệ thống trích xuất dữ liệu Nhựa Tiền Phong")
st.info("Phiên bản cập nhật: Tối ưu hóa nhận diện địa chỉ Bình Dương & TP.HCM")

uploaded_file = st.file_uploader("Upload file PDF scan của bạn", type="pdf")

if uploaded_file:
    with st.spinner('Đang quét dữ liệu từng trang...'):
        pdf_data = uploaded_file.read()
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        results = []
        stt_counter = 1
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            # Lấy text và chuẩn hóa: xóa khoảng trắng dư thừa để dễ quét Regex
            text = page.get_text("text")
            text_clean = " ".join(text.split())
            text_upper = text_clean.upper()
            
            # ĐIỀU KIỆN LỌC TRANG (Dựa trên thông tin bạn cung cấp)
            keywords = [
                "NHỰA THIẾU NIÊN TIỀN PHONG",
                "ĐỒNG AN 2",
                "HÒA PHÚ",
                "TP.TDM",
                "XÔ VIẾT NGHỆ TĨNH"
            ]
            
            if any(key in text_upper for key in keywords):
                # 1. Trích xuất NGÀY
                date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text_clean)
                ngay = date_match.group(1) if date_match else ""

                # 2. Trích xuất PR/SO
                # Quét các cụm bắt đầu bằng PR hoặc SO, lấy cả chuỗi phía sau cho đến khi gặp khoảng trắng
                pr_matches = re.findall(r"PR[\s]?[\d\.]+", text_upper)
                so_matches = re.findall(r"SO[\s]?[\d\.]+", text_upper)
                
                # Làm sạch và kết hợp
                pr_val = pr_matches[0].replace(" ", "") if pr_matches else ""
                so_val = so_matches[0].replace(" ", "") if so_matches else ""
                
                if pr_val and so_val:
                    pr_so_final = f"{pr_val}/{so_val}"
                else:
                    pr_so_final = pr_val if pr_val else so_val

                # 3. Trích xuất SM
                sm_match = re.search(r"SM[\s]?([\d\.,]+)", text_upper)
                sm_val = f"SM{sm_match.group(1).strip()}" if sm_match else ""

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
            st.success(f"Đã tìm thấy {len(df)} dòng dữ liệu!")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
            
            st.download_button(
                label="📥 Tải file Excel",
                data=output.getvalue(),
                file_name=f"KetQua_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Vẫn không tìm thấy dữ liệu. Có khả năng file này là 'PDF Image' hoàn toàn.")
            st.warning("Mẹo: Nếu bạn bôi đen được nhưng kết quả trắng, hãy thử mở PDF bằng Chrome, chọn In -> Lưu dưới dạng PDF (Save as PDF) rồi upload lại file mới đó lên đây.")
