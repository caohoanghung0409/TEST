import streamlit as st
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
import re
from paddleocr import PaddleOCR
import tempfile
import os

# ======================
# INIT OCR
# ======================
ocr = PaddleOCR(use_angle_cls=True, lang='vi')

st.set_page_config(page_title="PDF OCR CLEAN", layout="wide")
st.title("📄 PDF OCR -> Excel (Clean Version)")

# ======================
# CLEAN FUNCTION
# ======================
def extract_date(text):
    if pd.isna(text):
        return None
    m = re.search(r'\d{2}/\d{2}/\d{4}', str(text))
    return m.group(0) if m else None


def extract_pr(text):
    if pd.isna(text):
        return None
    m = re.search(r'PR\d{4}\.\d+', str(text))
    return m.group(0) if m else None


def is_noise(row):
    raw = " ".join([str(x) for x in row.values if pd.notna(x)])

    if len(raw.strip()) < 8:
        return True

    blacklist = [
        "phiếu", "giao hàng", "địa chỉ", "điện thoại",
        "fax", "mst", "dai dien", "ben giao", "ben nhan"
    ]

    if any(x in raw.lower() for x in blacklist):
        return True

    if re.search(r'PR\d{4}\.\d+', raw):
        return False

    if re.match(r'^\d+$', str(row.get("STT", ""))):
        return False

    return False


def clean_df(df):
    if "Ngày" in df.columns:
        df["Ngày"] = df["Ngày"].apply(extract_date)

    if "PR" in df.columns:
        df["PR"] = df["PR"].apply(extract_pr)

    df = df[~df.apply(is_noise, axis=1)]
    return df.reset_index(drop=True)

# ======================
# PDF TO IMAGE
# ======================
def pdf_to_images(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    images = []

    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = np.frombuffer(pix.tobytes(), dtype=np.uint8)
        img = cv2.imdecode(img, cv2.IMREAD_COLOR)
        images.append(img)

    return images

# ======================
# OCR
# ======================
def run_ocr(images):
    all_rows = []

    for img in images:
        result = ocr.ocr(img, cls=True)

        for line in result[0]:
            text = line[1][0]
            all_rows.append([text])

    df = pd.DataFrame(all_rows, columns=["Raw"])

    # ===== parse cơ bản =====
    df["STT"] = df["Raw"].apply(lambda x: re.findall(r'^\d+', str(x)))
    df["STT"] = df["STT"].apply(lambda x: x[0] if x else None)

    df["PR"] = df["Raw"].apply(extract_pr)
    df["Ngày"] = df["Raw"].apply(extract_date)

    return df

# ======================
# UI
# ======================
uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Đang xử lý OCR..."):
        images = pdf_to_images(uploaded_file)
        df = run_ocr(images)
        df = clean_df(df)

    st.success("Done!")

    st.subheader("📊 Data sau khi clean")
    st.dataframe(df)

    # ======================
    # EXPORT EXCEL
    # ======================
    output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    df.to_excel(output_path, index=False)

    with open(output_path, "rb") as f:
        st.download_button(
            "📥 Download Excel",
            f,
            file_name="ket_qua.xlsx"
        )
