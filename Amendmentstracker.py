import win32com.client
import pandas as pd
import os


outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6)
inbox2 = outlook.Folders('INBOX NAME').Folders('Sent Items').Folders('Amendments')

messages = inbox2.Items
message = messages.GetFirst()
rec_time = message.CreationTime
body_content = message.Body
subj_line = message.Subject
email_date = message.CreationTime
sender = message.Sender #does not read meeting invites - use messagesender()


EMnamesfile = open("\Namelist.txt", "r")
EMnameslist = []

datevar = "2020-11-18" ##STOP DATE GOES HERE

for x in EMnamesfile:
    x.replace('\n','')
    EMnameslist.append(x)


def messagesender():
    try:
        mes2 =message.Sender
        return message.Sender
    except AttributeError:
        pass

def messagebody():
    try:
        mes2 =message.Body
        return message.Body
    except AttributeError:
        pass

    
subjectlist = []

for x in range(len(messages)):
    emaildate = f"{message.CreationTime}"
    subject = "f{message.Subject}"
    body = f"{message.Body}"
    subjectlist.append(f"""
\n-------TIME: ------------\n
{message.CreationTime} \n
{message.CC} \n
-------- SUBJECT: --------\n
{message.Subject} \n
-------- BODY:---- \n
{body}""")
    if datevar in emaildate:  ##STOP DATE GOES HERE
        break
    message = messages.GetNext()


namelist = []

for entry in subjectlist:
    varentry = entry.lower().strip()
    for name in EMnameslist:
        varname = name.lower().strip()
        if varname in varentry:
            namelist.append(name.replace("\n",""))
                  

df = pd.DataFrame(namelist,  columns = [f"Amendments from today to {datevar}"])
df1 = df[f'S&S Amendments from today to {datevar}'].value_counts()
print(df1)

os.chdir(r"FILE LOCATION")
writer = pd.ExcelWriter(f'Amendments - Present to {datevar}.xlsx')
df1.to_excel(writer,'Sheet1')
writer.save()

        
