from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
import time

app = Application(backend="uia").start(
    r"C:\Program Files\Google\Chrome\Application\chrome.exe"
)

time.sleep(5)

chrome = app.window(title_re=".*Chrome.*")
chrome.set_focus()
chrome.maximize()

rect = chrome.rectangle()

# Click address bar (mouse)
mouse.click(coords=(rect.left + 300, rect.top + 60))
send_keys("https://mail.google.com{ENTER}")

time.sleep(10)

# Click Compose (mouse — approximate)
mouse.click(coords=(rect.left + 100, rect.top + 200))
time.sleep(2)

send_keys("receiver@gmail.com{TAB}")
send_keys("Mouse Automation Test{TAB}")
send_keys("This mail was sent using mouse + pywinauto")

send_keys("^enter")
