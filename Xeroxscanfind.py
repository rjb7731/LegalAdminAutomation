import win32com.client


outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6)
inbox = outlook.Folders('EMAIL HERE').Folders('Inbox')

messages = inbox.Items
message = messages.GetLast()
rec_time = message.CreationTime
body_content = message.body
subj_line = message.subject
sender = message.Sender #does not read meeting invites - use messagesender()

##while message:
##	print(message.subject)
##	message=messages.GetPrevious()


def messagesender():
    mes1 = message.subject
    try:
        mes2 =message.Sender
        return message.Sender
    except AttributeError:
        pass

def scannedimages():
    if "Scanned from a Xerox Multifunction Device" in message.subject:
        print(f"Email: {x}- Scanned image from- EMAIL")
        
for x in range(len(messages)):
    messagesender()
    #print(f"SENDER: {messagesender()}, SUBJECT: {message.subject}")
    if "692" in message.body:
        print(f"FOUND {message.body}")
    message = messages.GetPrevious()
