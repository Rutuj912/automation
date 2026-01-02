from pywinauto.application import Application
import os
from pathlib import Path

app= Application(backend="uia").start("notepad.exe").connect(title="Untitled - Notepad",timeout=100)

# app.UntitledNotepad.print_control_identifiers()

texteditor= app.UntitledNotepad.child_window(title="Text Editor", auto_id="15", control_type="Edit").wrapper_object()
texteditor.type_keys("my name is rutuj",with_spaces=True)

filemenu= app.UntitleNotepad.child_window(title="File", control_type="MenuItem").wrapper_object()
filemenu.click_input()

#app.UntitledNotepad.print_control_identifiers()

savefile= app.UntitledNotepad.child_window(title="Save As...	Ctrl+Shift+S", auto_id="4", control_type="MenuItem").wrapper_object()
savefile.click_input()

#app.UntitledNotepad.print_control_identifiers()

filename= app.UntitledNotepad. child_window(title="File name:", auto_id="1001", control_type="Edit").wrapper_object()

filename.type_keys("test.txt",with_spaces=True)

#app.UntitledNotepad.print_control_identifiers()

finalsave= app.UntitledNotepad.child_window(title="Save", auto_id="1", control_type="Button").wrapper_object()
finalsave.click_input()


filepath=r"C:\Users\rutuj.patel\Documents\test.txt"


if os.path.exists(filepath):
    replace= app.UntitledNotepad.child_window(title="Yes", auto_id="CommandButton_6", control_type="Button")
    replace.click_input()



