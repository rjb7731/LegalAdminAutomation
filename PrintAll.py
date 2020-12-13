import win32api
import win32print
import os
import time


file = os.listdir(r'FILE DIR HERE') #File path goes here

win32api.MessageBox(0, 'Did you change the printer settings to colour?', 'Printer Settings')

for x in file:
    win32api.ShellExecute(0,"printto",f'{x}','"%s"' % win32print.GetDefaultPrinter(),r"C:\Users\BradRy\OneDrive - The Co-op Group\Desktop\PRINTALL",0)
    print(x + '- SENT TO THE PRINTER')
    time.sleep(2)

print('------Completed------')
