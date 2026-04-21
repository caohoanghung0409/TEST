import streamlit as st
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="Lọc Phiếu Tiền Phong", page_icon="🚚")

st.title("🚚 Hệ thống Lọc Phiếu Nhựa Tiền Phong")
st.info("Hệ thống sẽ tự động giữ lại các trang là PHIẾU GIAO HÀNG của NHỰA TIỀN PHONG.")

uploaded_file = st.file_uploader("Kéo thả file PDF vào đây", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý...'):
        # Đọc dữ liệu
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        out_doc = fitz.open()
        
        found_pages = []
        for i in range(len(doc)):
            page = doc.load_page(i)
            text = page.get_text().upper()
            
            if "PHIẾU GIAO HÀNG" in text and "NHỰA TIỀN PHONG" in text:
                out_doc.insert_pdf(doc, from_page=i, to_page=i)
                found_pages.append(i + 1)

        if found_pages:
            st.success(f"Đã tìm thấy {len(found_pages)} trang hợp lệ!")
            st.write(f"Số trang: {found_pages}")
            
            # Xuất file
            pdf_bytes = out_doc.tobytes()
            st.download_button(
                label="📥 Tải xuống kết quả",
                data=pdf_bytes,
                file_name=f"Da_Loc_{uploaded_file.name}",
                mime="application/pdf"
            )
        else:
            st.error("Không tìm thấy trang nào chứa đúng mẫu yêu cầu.")
        
        doc.close()
        out_doc.close()
