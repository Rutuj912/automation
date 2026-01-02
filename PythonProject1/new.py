import fitz  # PyMuPDF
import easyocr
import numpy as np
import json
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import time
import os

# ----------------------------
# STEP 1: Extract text from PDF
# ----------------------------
pdf_path = "sample.pdf"  # Replace with your PDF path
reader = easyocr.Reader(['en'])

all_text = []

pdf_doc = fitz.open(pdf_path)

for page_num in range(len(pdf_doc)):
    page = pdf_doc[page_num]
    pix = page.get_pixmap(dpi=300)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    results = reader.readtext(img, detail=0)
    all_text.extend(results)

# Combine all extracted text into a single string
ocr_text = "\n".join(all_text)

# Optional: Save OCR to JSON
ocr_json = {"extracted_text": all_text}
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(ocr_json, f, indent=4, ensure_ascii=False)

print("OCR extraction done. JSON saved to output.json")

# ----------------------------
# STEP 2: Open Word & Paste Text
# ----------------------------
app = Application(backend="uia").start(
    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
)
time.sleep(3)

word_window = app.connect(title_re=".*Word").window(title_re=".*Word")

# Open Blank Document
try:
    open_document = word_window.child_window(title="Blank document", control_type="ListItem")
    open_document.click_input()
    time.sleep(1)
except:
    pass

# Type OCR text into Word using send_keys
# Send line by line to handle newlines
for line in ocr_text.splitlines():
    send_keys(line)
    send_keys('{ENTER}')
time.sleep(1)

# ----------------------------
# STEP 3: Save Word File
# ----------------------------
file_tab = word_window.child_window(title="File Tab", auto_id="FileTabButton", control_type="Button")
file_tab.click_input()
time.sleep(1)

save_file = word_window.child_window(title="Save As", control_type="ListItem")
save_file.click_input()
time.sleep(1)

this_pc = word_window.child_window(title="This PC", control_type="TabItem")
this_pc.click_input()
time.sleep(1)

desktop = word_window.child_window(title="Desktop", control_type="ListItem")
desktop.click_input()
time.sleep(1)

filename = word_window.child_window(title="File name:", auto_id="FileNameControlHost", control_type="ComboBox")
filename.click_input()
send_keys("OCR_Result", with_spaces=True)
time.sleep(0.5)

filesave = word_window.child_window(title="Save", auto_id="1", control_type="Button")
filesave.click_input()
time.sleep(2)

# Handle "Replace existing file" dialog if it appears
try:
    replace_dialog = word_window.child_window(title="Microsoft Word", control_type="Window")
    save_different = replace_dialog.child_window(title="Save changes with a different name.", control_type="RadioButton")
    save_different.click_input()
    time.sleep(1)
    filename = replace_dialog.child_window(title="File name:", auto_id="FileNameControlHost", control_type="ComboBox")
    filename.click_input()
    send_keys("OCR_Result_v2")
    time.sleep(0.5)
    newsave = replace_dialog.child_window(title="Save", auto_id="1", control_type="Button")
    newsave.click_input()
    time.sleep(1)
except:
    pass

# Close Word safely
close_button = word_window.child_window(title="Close", control_type="Button", top_level_only=True)
close_button.click_input()
time.sleep(1)

# If Word asks "Do you want to save changes?", click "Don't Save"
try:
    dont_save = app.window(title_re=".*Word").child_window(title="Don't Save", control_type="Button")
    dont_save.click_input()
except:
    pass

print("OCR text successfully written to Word and saved.")
