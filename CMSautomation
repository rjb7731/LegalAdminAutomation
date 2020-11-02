from openpyxl import load_workbook
import pyautogui
import pygetwindow
import time
import pyperclip


window = pygetwindow.getWindowsWithTitle('WINDOW TITLE')[0]

if window.isMaximized == False:
    window.maximize()

workbook = load_workbook(filename="Spreadsheets\\Matters.xlsx")
sheet = workbook.active

FirstRow = []
SecondRow = []
Mattername = []

matnum = str(sheet.max_row)

time.sleep(2)

for row in sheet['A1':'A'+matnum]:  # Set First Entry from spreadsheet
    for cell in row:
        FirstRow.append(cell.value)

for row in sheet['B1':'B'+matnum]:  # Set Second Data Entry from spreadsheet
    for cell in row:
        SecondRow.append(cell.value)

for row in sheet['C1':'C'+matnum]:  # Set Second Data Entry from spreadsheet
    for cell in row:
        Mattername.append(cell.value)

def loadmatter():
    """enter entries and load matter"""
    pyautogui.moveTo(408, 199)  # move to first position
    pyautogui.click()
    pyautogui.typewrite(str(FirstRow[entry]))
    pyautogui.moveTo(707, 200)  # move to second position
    pyautogui.click()
    pyautogui.typewrite(str(SecondRow[entry]))
    pyautogui.moveTo(1258, 209)  # moves to load button
    pyautogui.click()

def datecheck():
    """check date"""
    pyautogui.moveTo(1250,349) 
    pyautogui.click()

def fire():
    """destroy file"""
    pyautogui.moveTo(1544, 258)  # moves to fire button
    pyautogui.click()
    pyautogui.moveTo(979, 581)  # moves to yes button
    time.sleep(1)
    pyautogui.click()
    time.sleep(3)

def clearandback():
    """move back to start and clear entries"""
    pyautogui.moveTo(408, 199)  # moves back to first box to clear
    time.sleep(0.5)
    pyautogui.doubleClick()
    pyautogui.press('backspace')
    pyautogui.moveTo(707, 200)  # moves to second box to clear
    time.sleep(0.5)
    pyautogui.doubleClick()
    pyautogui.press('backspace')

def checkdescription():
    """move to description to check"""
    pyautogui.moveTo(1175, 393) 
    pyautogui.click()
    pyautogui.keyDown('shift')
    time.sleep(0.1)
    pyautogui.press('end')
    time.sleep(0.1)
    pyautogui.keyUp('shift')
    pyperclip.copy()
    found = pyperclip.paste()
    sheetfound = f'{Mattername[entry]}'
    if found == sheetfound:
        pyautogui.alert("Match")
    else:
        pyautogui.alert(f"ERROR: These Do Not Match, Sheet = {Mattername[entry]}, CMS = {sheetfound}")
        
def copy_clipboard():
    """copy to clipbaord"""
    pyperclip.copy("")# <- This empties the clipboard
    pyautogui.moveTo(1175, 393) #move to description to check
    pyautogui.click()
    pyautogui.keyDown('shift')
    time.sleep(0.1)
    pyautogui.press('end')
    time.sleep(0.1)
    pyautogui.keyUp('shift')
    pyperclip.copy()
    time.sleep(.01) 
    return pyperclip.paste()

for entry in range(1, int(matnum)):  # set iterations
    loadmatter()
    pyautogui.alert(text=f'{Mattername[entry]}', title='Move to Next?', button='Next')
    clearandback()
    
