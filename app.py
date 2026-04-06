import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
import tempfile
import zipfile
import os
import time
from openpyxl import load_workbook

# =========================
# USER LOGIN
# =========================
USERS = {
    "user1": "123",
    "user2": "456"
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

# =========================
# LOGIN UI
# =========================
if not st.session_state.logged_in:

    st.set_page_config(page_title="Login", layout="centered")

    st.markdown("## 🔐 Đăng nhập")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success("✅ Login thành công")
            st.rerun()
        else:
            st.error("❌ Sai tài khoản hoặc mật khẩu")

    st.stop()

# =========================
# CONFIG APP
# =========================
st.set_page_config(page_title="OCR Drive UI", layout="wide")

# =========================
# SESSION
# =========================
if "processing" not in st.session_state:
    st.session_state.processing = False

if "done" not in st.session_state:
    st.session_state.done = False

if "clear_uploader" not in st.session_state:
    st.session_state.clear_uploader = False

# =========================
# STYLE
# =========================
st.markdown("""
<style>
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.stApp { background: #f8fafc; }

.header {
    display:flex;
    justify-content:space-between;
    align-items:center;
    font-size:20px;
    font-weight:600;
    padding:10px 0;
}

.file-row {
    font-size:14px;
    margin-top:10px;
}

.progress {
    height:6px;
    background:#e5e7eb;
    border-radius:10px;
    overflow:hidden;
}

.progress-bar {
    height:100%;
    background:linear-gradient(90deg,#0ea5e9,#22c55e);
}

/* GLOBAL BAR */
.global-bar {
    position:relative;
    height:18px;
    background:#e5e7eb;
    border-radius:999px;
    overflow:hidden;
}

.global-fill {
    height:100%;
    background:linear-gradient(90deg,#6366f1,#22c55e);
}

.global-text {
    position:absolute;
    width:100%;
    text-align:center;
    font-size:12px;
    top:0;
    line-height:18px;
    font-weight:600;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER + LOGOUT
# =========================
col1, col2 = st.columns([6,1])

with col1:
    st.markdown(f'<div class="header">📁 OCR TOOL | 👤 {st.session_state.username}</div>', unsafe_allow_html=True)

with col2:
    if st.button("🚪 Logout"):
        st.session_state.logged_in = False
        st.rerun()

# =========================
# UPLOADER
# =========================
uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# GLOBAL BAR
# =========================
def render_bar(p):
    return f"""
<div style="margin:10px 0;">
    <div class="global-bar">
        <div class="global-fill" style="width:{p}%"></div>
        <div class="global-text">{p}%</div>
    </div>
</div>
"""

# =========================
# PROCESS
# =========================
def extract(file, box, global_box, done, total):
    results = []
    imgs = convert_from_bytes(file.read(), dpi=150)

    for i, img in enumerate(imgs, 1):
        done[0] += 1
        percent = int((done[0] / total) * 100)

        global_box.markdown(render_bar(percent), unsafe_allow_html=True)

        box.markdown(f"""
<div class="file-row">
📄 {file.name} • Trang {i}/{len(imgs)}
<div class="progress"><div class="progress-bar" style="width:{int(i/len(imgs)*100)}%"></div></div>
</div>
""", unsafe_allow_html=True)

        w,h = img.size
        img = img.crop((0,0,w,int(h*0.4)))

        sm,date = process_page(img)
        if sm and date:
            results.append({"SM":sm,"Ngày":date})

    return results

# =========================
# MAIN
# =========================
if uploaded_files:

    total_pages = sum(len(convert_from_bytes(f.read(), dpi=50)) for f in uploaded_files)
    for f in uploaded_files:
        f.seek(0)

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Process"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        done = [0]
        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            for i,f in enumerate(uploaded_files):

                data = extract(f, boxes[i], global_box, done, total_pages)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0,"STT",range(1,len(df)+1))

                    name = os.path.splitext(f.name)+".xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name,index=False)

                        wb = load_workbook(tmp.name)
                        ws = wb.active

                        for col in ws.columns:
                            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                            ws.column_dimensions[col[0].column_letter].width = max_len + 3

                        wb.save(tmp.name)
                        zipf.write(tmp.name, os.path.basename(name))

        st.session_state.zip = zip_buffer.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DOWNLOAD
# =========================
if st.session_state.done:

    st.success("🎉 Xong!")

    with open(st.session_state.zip,"rb") as f:
        zip_data = f.read()

    if st.download_button("📥 Download ZIP", zip_data, "ocr_results.zip"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
