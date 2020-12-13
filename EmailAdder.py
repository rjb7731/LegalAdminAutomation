import time
import pyautogui

for x in range(0, 20):
    screenWidth, ScreenHeight = pyautogui.size()
    CurrentMouseX, CurrentMouseY = pyautogui.position()
    pyautogui.moveTo(1190, 324)  #irst OK Box
    time.sleep(2)
    pyautogui.click()
    pyautogui.moveTo(1480, 503) # Attachment box
    time.sleep(8)
    pyautogui.click()

print("Adding Complete")
