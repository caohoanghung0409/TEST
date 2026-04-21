import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="Fix Dứt Điểm Tiền Phong", layout="wide")
st.title("✅ Trích xuất Nhựa Tiền Phong - Bản Final")

uploaded_file = st.file_uploader("Upload file PDF", type="pdf")

def clean_text(t):
    if not t: return ""
    # Gom các ký tự rời rạc lại và xóa khoảng trắng thừa
    return " ".join(t.split())

if uploaded_file:
    with st.spinner('Đang dùng thuật toán quét sâu...'):
        results = []
        stt = 1
        
        with pdfplumber.open(uploaded_file) as pdf:
            for i, page in enumerate(pdf.pages):
                # Trích xuất văn bản theo cụm (giúp trị lỗi chữ rời rạc)
                text = page.extract_text()
                if not text: continue
                
                text_upper = text.upper()
                
                # Điều kiện: Chỉ cần có mã PR hoặc SO hoặc SM hoặc Tiền Phong
                # Vì file scan lỗi nên mình ưu tiên tìm mã số trước
                has_identity = any(kw in text_upper for kw in ["PR26", "SO26", "SM26", "TIỀN PHONG", "TIEN PHONG"])
                
                if has_identity:
                    # 1. Tìm Ngày
                    date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
                    ngay = date_match.group(1) if date_match else ""

                    # 2. Tìm PR/SO (Dùng regex nới lỏng cho file scan)
                    pr = re.search(r"PR\s?(\d{4}[\d\.]*)", text_upper)
                    so = re.search(r"SO\s?(\d{4}[\d\.]*)", text_upper)
                    
                    p_val = f"PR{pr.group(1)}" if pr else ""
                    s_val = f"SO{so.group(1)}" if so else ""
                    
                    pr_so = f"{p_val}/{s_val}" if (p_val and s_val) else (p_val or s_val)

                    # 3. Tìm SM
                    sm = re.search(r"SM\s?(\d{4}[\d\.,]*)", text_upper)
                    sm_val = f"SM{sm.group(1).replace(',', '.')}" if sm else ""

                    if pr_so or sm_val:
                        results.append({
                            "STT": stt,
                            "SM": sm_val,
                            "PR/SO": pr_so,
                            "NGÀY": ngay,
                            "SỐ TRANG": i + 1
                        })
                        stt += 1

        if results:
            df = pd.DataFrame(results)
            st.success(f"Đã tìm thấy {len(df)} dòng dữ liệu!")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
            
            st.download_button(label="📥 Tải Excel", data=output.getvalue(), 
                             file_name="KetQua_Final.xlsx", mime="application/vnd.ms-excel")
        else:
            st.error("Không tìm thấy dữ liệu. File scan này có lớp chữ bị lỗi hoàn toàn.")
            st.info("Mẹo: Nếu vẫn lỗi, hãy dùng Chrome mở file -> In -> Lưu PDF. Cách này chắc chắn 100% sẽ tạo lại lớp chữ sạch.")
