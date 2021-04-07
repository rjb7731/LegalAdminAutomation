import win32com.client
import datetime
import os
import pyautogui
from PyPDF2 import PdfFileMerger

merger = PdfFileMerger(strict=False)

os.startfile(r"DESIGNATED FOLDER")
os.chdir(r"DESIGNATED FOLDER")


outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6)
inbox = outlook.Folders('EMAIL').Folders('Inbox')

messages = inbox.Items
message = messages.GetLast()
rec_time = message.CreationTime
body_content = message.body
subj_line = message.subject


time_now = f"{datetime.datetime.now()}"


def messagesender():
    mes1 = message.subject
    try:
        mes2 =message.Sender
        return message.Sender
    except AttributeError:
        pass
                             

input_no = pyautogui.prompt(text='How many scanned pdf emails are there?', title='Email_pdf_to_file' , default='')
input_added = int(input_no)-1  ## to account for '0'

num = 1

for x in range(len(messages)):
    email_date = f"{message.CreationTime}"
    if f"{time_now[0:10]}" in email_date:
        pass
    else:
        break
    if "Scanned from a Xerox" in message.subject:
        Item_Attach = message.Attachments
        for item in Item_Attach:
            attachment = Item_Attach.item(1)
            attachment.SaveAsFile(r'DESIGNATED FOLDER\\'+f"{num}.pdf")
            merger.append(r'DESIGNATED FOLDER\\'+f"{num}.pdf")
        num += 1
    if len(os.listdir(r"DESIGNATED FOLDER")) > int(input_added):
           break
    message = messages.GetPrevious()
    
merger.write('Completed Merge.pdf')       
merger.close()



    
