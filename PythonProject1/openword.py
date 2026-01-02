# import fitz  # PyMuPDF
# import easyocr
# import numpy as np
# import cv2
# import json
#
# # Path to PDF
# pdf_path = "sample.pdf"
#
# # Initialize EasyOCR reader
# reader = easyocr.Reader(['en'])
#
# all_text = []
#
# # Open PDF
# pdf_doc = fitz.open(pdf_path)
#
# for page_num in range(len(pdf_doc)):
#     page = pdf_doc[page_num]
#
#     # Render page as image (pix)
#     pix = page.get_pixmap(dpi=300)  # dpi=300 for better OCR
#
#     # Convert to OpenCV format
#     img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
#
#     # OCR
#     results = reader.readtext(img, detail=0)
#
#     all_text.extend(results)
#
# # Convert to JSON
# ocr_json = {"extracted_text": all_text}
#
# # Print JSON
# print(json.dumps(ocr_json, indent=4, ensure_ascii=False))
#
# # Save JSON to file
# with open("output.json", "w", encoding="utf-8") as f:
#     json.dump(ocr_json, f, indent=4, ensure_ascii=False)
#
# print("OCR text saved to output.json")
#
# #
# from pywinauto.application import Application
# import time
# import os
# from pywinauto.keyboard import send_keys
#
#
# app = Application(backend="uia").start(
#     r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
# )
#
# time.sleep(3)
#
# word_window = app.connect(title_re=".*Word").window(title_re=".*Word")
#
# open_document= word_window.child_window(title="Blank document", control_type="ListItem")
# open_document.click_input()
#
# ocr_text = ocr_json["extracted_text"]  # This is a list
#
# for line in ocr_text:
#     word_window.type_keys(line, with_spaces=True)
#     word_window.type_keys("{ENTER}")  # Move to next line
#
#
# file_tab= word_window.child_window(title="File Tab", auto_id="FileTabButton", control_type="Button")
# file_tab.click_input()
#
# save_file= word_window.child_window(title="Save As", control_type="ListItem")
# save_file.click_input()
#
# this_pc = word_window.child_window(title="This PC", control_type="TabItem")
# this_pc.click_input()
#
# desktop= word_window.child_window(title="Desktop", control_type="ListItem")
# desktop.click_input()
#
# filename= word_window.child_window(title="File name:", auto_id="FileNameControlHost", control_type="ComboBox")
# filename.click_input()
# filename.type_keys("test1",with_spaces=True)
#
# filesave= word_window.child_window(title="Save", auto_id="1", control_type="Button")
# filesave.click_input()
#
# close= word_window.child_window(title="Close", control_type="Button",found_index=0)
# close.click_input()
#
#
# filepath=(r"C:\Users\rutuj.patel\Desktop")
#
# if os.path.exists(filepath):
#     savewithdifferent= word_window.child_window(title="Save changes with a different name.", control_type="RadioButton")
#     savewithdifferent.click_input()
#
#     clickok= word_window.child_window(title="OK", control_type="Button")
#     clickok.click_input()
#
#     editname= word_window.child_window(title="File name:", auto_id="FileNameControlHost", control_type="ComboBox")
#     editname.type_keys("test2",with_spaces=True)
#
#     newsave= word_window.child_window(title="Save", auto_id="1", control_type="Button")
#     newsave.click_input()
#
# close= word_window.child_window(title="Close", control_type="Button",found_index=0)
# close.click_input()
#
#
# # #===============================================================================EMAIL PART================================================================================
from pywinauto.application import Application
import time


outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
app = Application(backend="uia").start(outlook_path)
time.sleep(15)


outlook_window = app.window(title_re=".*Outlook.*")
outlook_window.wait("visible", timeout=30)


new_email_button = outlook_window.child_window(title="New Email", auto_id="NewItem", control_type="Button")
new_email_button.click_input()

email_window = app.window(title_re=".*- Message .*")
email_window.wait("visible", timeout=20)  # waits until new window is visible


# Focus first Edit control (To field)
to_field = email_window.child_window(control_type="Edit", found_index=0)
to_field.set_focus()
to_field.type_keys("spandan@zylitix.ai", with_spaces=True)
to_field.click_input()


subject= email_window.child_window(title="Subject", auto_id="4516", control_type="Text")
subject.set_focus()
subject.type_keys("Automation Task – PDF Data Extraction and Email Sharing ", with_spaces=True)


message= email_window.child_window(title="Message", auto_id="547", control_type="Text")
message.set_focus()
message.type_keys("Hello, here is the word file attached of the extracted document", with_spaces=True)


attach= email_window.child_window(title="Attachment", auto_id="547")
attach.set_focus()
attach.click_input()

email_window.print_control_identifiers()
