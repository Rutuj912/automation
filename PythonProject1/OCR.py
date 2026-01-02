import fitz  # PyMuPDF
import easyocr
import numpy as np
import cv2
import json

# Path to PDF
pdf_path = "sample.pdf"

# Initialize EasyOCR reader
reader = easyocr.Reader(['en'])

all_text = []

# Open PDF
pdf_doc = fitz.open(pdf_path)

for page_num in range(len(pdf_doc)):
    page = pdf_doc[page_num]

    # Render page as image (pix)
    pix = page.get_pixmap(dpi=300)  # dpi=300 for better OCR

    # Convert to OpenCV format
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

    # OCR
    results = reader.readtext(img, detail=0)

    all_text.extend(results)

# Convert to JSON
ocr_json = {"extracted_text": all_text}

# Print JSON
print(json.dumps(ocr_json, indent=4, ensure_ascii=False))

# Save JSON to file
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(ocr_json, f, indent=4, ensure_ascii=False)

print("OCR text saved to output.json")
