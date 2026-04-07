```python
import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import os
import time
import base64
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="THL PDF TO EXCEL", layout="wide")

# =========================
# SESSION
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False

if "done" not in st.session_state:
    st.session_state.done = False

if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

if "last_uploaded_names" not in st.session_state:
    st.session_state.last_uploaded_names = []

if "excel_files" not in st.session_state:
    st.session_state.excel_files = []

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 0.5rem !important;}
.stApp { background: #f1f5f9; }

.header {
    font-size:22px;
    font-weight:700;
    margin-bottom:10px;
}

div.stButton > button {
    background: linear-gradient(135deg,#3b82f6,#22c55e);
    color:white;
    border:none;
    border-radius:12px;
    padding:14px 28px;
    font-weight:600;
    font-size:16px;
    box-shadow:0 4px 14px rgba(0,0,0,0.15);
    transition: all 0.25s ease;
}

div.stButton > button:hover {
    transform: translateY(-2px) scale(1.02);
    box-shadow:0 8px 20px rgba(0,0,0,0.2);
}

.new-btn button {
    background: linear-gradient(135deg,#f59e0b,#ef4444) !important;
}

.done-box {
    display:flex;
    justify-content:center;
    align-items:center;
    height:80px;
    font-size:26px;
    font-weight:800;
    color:#16a34a;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 THL PDF → EXCEL </div>', unsafe_allow_html=True)

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "📂 Chọn file PDF",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

current_names = [f.name for f in uploaded_files] if uploaded_files else []

if current_names != st.session_state.last_uploaded_names:
    st.session_state.processing = False
    st.session_state.done = False
    st.session_state.last_uploaded_names = current_names

# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(
        img,
        lang='eng',
        config='--oem 3 --psm 6'
    )
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# PROCESS PDF
# =========================
def extract_pdf(file):
    results = []
    images = convert_from_bytes(file.read(), dpi=150)

    for img in images:
        w, h = img.size
        img = img.crop((0, 0, w, int(h * 0.4)))

        sm, date = process_page(img)
        if sm and date:
            results.append({
                "SM": sm,
                "Ngày": date
            })

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        st.write("⏳ Đang xử lý...")

        excel_paths = []

        for f in uploaded_files:
            data = extract_pdf(f)

            if data:
                df = pd.DataFrame(data)
                df.insert(0, "STT", range(1, len(df)+1))

                name = os.path.splitext(f.name)[0] + ".xlsx"

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    df.to_excel(tmp.name, index=False)

                    wb = load_workbook(tmp.name)
                    ws = wb.active

                    for col in ws.columns:
                        max_len = max(
                            len(str(c.value)) if c.value else 0
                            for c in col
                        )
                        ws.column_dimensions[col[0].column_letter].width = max_len + 3

                    wb.save(tmp.name)

                    excel_paths.append((name, tmp.name))

        st.session_state.excel_files = excel_paths
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD (KHÔNG ZIP)
# =========================
if st.session_state.done:

    st.markdown(
        '<div class="done-box">🎉 HOÀN THÀNH !!!</div>',
        unsafe_allow_html=True
    )

    # AUTO DOWNLOAD từng file
    for name, path in st.session_state.excel_files:
        with open(path, "rb") as f:
            data = f.read()

        b64 = base64.b64encode(data).decode()

        st.markdown(f"""
        <iframe src="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
        style="display:none;"></iframe>
        """, unsafe_allow_html=True)

    # BUTTON DOWNLOAD THỦ CÔNG (backup)
    for name, path in st.session_state.excel_files:
        with open(path, "rb") as f:
            st.download_button(
                label=f"⬇️ Tải {name}",
                data=f,
                file_name=name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # RESET
    st.markdown('<div class="new-btn">', unsafe_allow_html=True)
    if st.button("🔄 XỬ LÝ FILE MỚI"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.session_state.excel_files = []
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
```
