import win32com.client
import datetime
import re
import os

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6)
inbox = outlook.Folders(INBOX NAME).Folders('Inbox').Folders('SUBFOLDER')
inbox2 = outlook.Folders('INBOXNAME)').Folders('Inbox').Folders('SUBFOLDER').Folders('INSIDE SUBFOLDER')

folder_names = inbox.Folders
email_folder = inbox2.Folders

for x in email_folder:
    os.chdir(r"FILE PATH TO SAVE GOES HERE")
    print(f"FOLDER = {x}")
    os.mkdir(f"FILE PATH TO MAKE GOES HERE \\{x}")
    os.chdir(f" NEW MADE FILE PATH GOES HERE \\{x}")
    indv_folder = x.Items
    for message in indv_folder:
        email = str(message.Subject)
        name = re.sub('[./"]+', '', email)
        print(name)
        email_time = str(message.CreationTime)
        try:
            message.saveAs(os.getcwd()+"\\"+name+"-"+email_time[0:10]+".msg")
        except AttributeError:
            pass
