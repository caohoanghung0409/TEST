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
# CONFIG
# =========================
st.set_page_config(page_title="OCR Ultimate UI", layout="wide")

# =========================
# SESSION
# =========================
for key, default in {
    "processing": False,
    "done": False,
    "clear_uploader": False
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# =========================
# DARK MODE DETECT
# =========================
is_dark = st.get_option("theme.base") == "dark"

# =========================
# STYLE ULTIMATE
# =========================
st.markdown(f"""
<style>
header, #MainMenu, footer {{visibility:hidden;}}

.stApp {{
    background: {"#0f172a" if is_dark else "#f8fafc"};
    color: {"#e2e8f0" if is_dark else "#0f172a"};
}}

.header {{
    font-size:22px;
    font-weight:700;
    padding:10px 0;
}}

[data-testid="stFileUploader"] {{
    border: 2px dashed #6366f1;
    padding: 40px;
    border-radius: 20px;
    background: {"#1e293b" if is_dark else "white"};
    text-align:center;
}}

.stats {{
    display:flex;
    gap:20px;
    margin:10px 0;
}}

.card {{
    flex:1;
    padding:12px;
    border-radius:12px;
    background: {"#1e293b" if is_dark else "white"};
    box-shadow:0 4px 10px rgba(0,0,0,0.05);
    text-align:center;
    font-size:14px;
}}

.global-bar {{
    position:relative;
    height:20px;
    border-radius:999px;
    overflow:hidden;
    background:#334155;
}}

.global-fill {{
    height:100%;
    transition:width .4s;
    background:linear-gradient(90deg,#6366f1,#22c55e);
}}

.global-fill::before {{
    content:"";
    position:absolute;
    width:100%;
    height:100%;
    background: repeating-linear-gradient(
        45deg,
        rgba(255,255,255,0.2) 0,
        rgba(255,255,255,0.2) 10px,
        transparent 10px,
        transparent 20px
    );
    animation: move 1s linear infinite;
}}

@keyframes move {{
    from {{background-position:0 0;}}
    to {{background-position:40px 0;}}
}}

.global-text {{
    position:absolute;
    width:100%;
    text-align:center;
    top:0;
    font-size:12px;
    line-height:20px;
    font-weight:600;
}}

.file-row {{
    margin-top:10px;
    font-size:14px;
}}

.progress {{
    height:6px;
    background:#334155;
    border-radius:10px;
    overflow:hidden;
}}

.progress-bar {{
    height:100%;
    background:linear-gradient(90deg,#0ea5e9,#22c55e);
}}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown('<div class="header">🚀 OCR PDF → Excel ULTIMATE</div>', unsafe_allow_html=True)

# =========================
# UPLOADER
# =========================
uploader_key = "uploader_1" if not st.session_state.clear_uploader else "uploader_2"

uploaded_files = st.file_uploader(
    "📂 Kéo thả file PDF vào đây",
    type=["pdf"],
    accept_multiple_files=True,
    key=uploader_key
)

# =========================
# OCR
# =========================
def process_page(img):
    text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
    sm = re.search(r"(SM\d{4}\.\d{4})", text)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    return (sm.group(1), date.group(1)) if sm and date else (None, None)

# =========================
# RENDER BAR
# =========================
def render_bar(p, speed, eta):
    return f"""
<div style="margin:10px 0;">
    <div style="display:flex;justify-content:space-between;font-size:12px;">
        <div>⚡ {p}%</div>
        <div>🚀 {speed:.2f} p/s • ⏳ {eta}s</div>
    </div>
    <div class="global-bar">
        <div class="global-fill" style="width:{p}%"></div>
        <div class="global-text">{p}%</div>
    </div>
</div>
"""

# =========================
# PROCESS
# =========================
def extract(file, box, global_box, start, done_pages, total_pages):
    results = []
    imgs = convert_from_bytes(file.read(), dpi=150)

    for i, img in enumerate(imgs, 1):
        done_pages[0] += 1

        percent = int((done_pages[0] / total_pages) * 100)
        elapsed = time.time() - start
        speed = done_pages[0]/elapsed if elapsed>0 else 0
        eta = int((total_pages - done_pages[0]) / speed) if speed>0 else 0

        global_box.markdown(render_bar(percent, speed, eta), unsafe_allow_html=True)

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

    st.markdown(f"""
<div class="stats">
<div class="card">📄 Files<br><b>{len(uploaded_files)}</b></div>
<div class="card">📑 Pages<br><b>{total_pages}</b></div>
</div>
""", unsafe_allow_html=True)

    global_box = st.empty()
    boxes = [st.empty() for _ in uploaded_files]

    if not st.session_state.processing and not st.session_state.done:
        if st.button("🚀 Bắt đầu xử lý"):
            st.session_state.processing = True
            st.rerun()

    if st.session_state.processing:

        start = time.time()
        done_pages = [0]

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            for i,f in enumerate(uploaded_files):

                data = extract(f, boxes[i], global_box, start, done_pages, total_pages)

                if data:
                    df = pd.DataFrame(data)
                    df.insert(0,"STT",range(1,len(df)+1))

                    name = os.path.splitext(f.name)[0]+".xlsx"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        df.to_excel(tmp.name,index=False)

                        wb = load_workbook(tmp.name)
                        ws = wb.active

                        for col in ws.columns:
                            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                            ws.column_dimensions[col[0].column_letter].width = max_len + 3

                        wb.save(tmp.name)
                        zipf.write(tmp.name,name)

        st.session_state.zip = zip_buffer.name
        st.session_state.processing = False
        st.session_state.done = True
        st.rerun()

# =========================
# DONE
# =========================
if st.session_state.done:

    st.success("🎉 Xong rồi!")

    # 🔔 SOUND
    st.markdown("""
    <audio autoplay>
    <source src="https://www.soundjay.com/buttons/sounds/button-3.mp3" type="audio/mp3">
    </audio>
    """, unsafe_allow_html=True)

    with open(st.session_state.zip,"rb") as f:
        zip_data = f.read()

    if st.download_button("📥 Download ZIP", zip_data, "ocr_results.zip"):
        st.session_state.done = False
        st.session_state.clear_uploader = not st.session_state.clear_uploader
        st.rerun()
