import fitz  # PyMuPDF
import cv2
import numpy as np
from paddleocr import PaddleOCR
import re

# ================= OCR INIT =================
ocr = PaddleOCR(use_angle_cls=True, lang='vi')

# ================= CONFIG =================
ANCHOR_KEYWORDS = [
    "NHUA THIEU NIEN TIEN PHONG",
    "NHUA TIEN PHONG",
]

PHIEU_KEYWORDS = [
    "PHIEU GIAO HANG"
]

# Regex chuẩn 3 case
PATTERN = re.compile(
    r"SO\s*:?\s*(?:(SM\d{4}\.\d{4})|(?:(PR\d{4}\.\d{4})\/(SO\d{4}\.\d{4}))|(SO\d{4}\.\d{4}))\s*NGAY\s*:?\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE
)

# ================= HELPER =================
def normalize_text(text):
    text = text.upper()
    text = text.replace(":", " ")
    text = re.sub(r"\s+", " ", text)
    return text

def extract_text_from_image(img):
    result = ocr.ocr(img, cls=True)
    lines = []
    for line in result:
        for word in line:
            lines.append(word[1][0])
    return lines

def rotate_image(image, angle):
    if angle == 0:
        return image
    elif angle == 180:
        return cv2.rotate(image, cv2.ROTATE_180)

# ================= CORE LOGIC =================
def process_page(image):
    results = []

    for angle in [0, 180]:
        img_rotated = rotate_image(image, angle)
        lines = extract_text_from_image(img_rotated)
        lines_norm = [normalize_text(l) for l in lines]

        # ===== 1. CHECK ANCHOR =====
        has_anchor = any(
            any(k in line for k in ANCHOR_KEYWORDS)
            for line in lines_norm
        )

        if not has_anchor:
            continue

        # ===== 2. FIND "PHIEU GIAO HANG" =====
        for i, line in enumerate(lines_norm):
            if any(k in line for k in PHIEU_KEYWORDS):

                # ===== 3. CHỈ LẤY 3-5 DÒNG NGAY SAU =====
                search_zone = lines_norm[i:i+6]

                for zone_line in search_zone:
                    match = PATTERN.search(zone_line)
                    if match:
                        sm = match.group(1)
                        pr = match.group(2)
                        so = match.group(3) or match.group(4)
                        date = match.group(5)

                        results.append({
                            "SM": sm if sm else "",
                            "PR": pr if pr else "",
                            "SO": so if so else "",
                            "Ngày": date
                        })

                        return results  # stop luôn khi tìm thấy

    return results

# ================= PDF PROCESS =================
def process_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    final_results = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap()
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        page_results = process_page(img)

        for r in page_results:
            r["Page"] = page_num + 1
            final_results.append(r)

    return final_results

# ================= RUN TEST =================
if __name__ == "__main__":
    pdf_path = "test.pdf"  # đổi file tại đây
    data = process_pdf(pdf_path)

    for row in data:
        print(row)
