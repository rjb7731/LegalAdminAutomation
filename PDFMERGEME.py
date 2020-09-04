import os
from PyPDF2 import PdfFileMerger

link = ##Folder link goes here

os.chdir(link)
 
source_dir = os.listdir(link)
list= []
list2=[]
merger = PdfFileMerger()
source = int(len(source_dir))

for entry in range(1, source):
    newentry = str(entry)+ ".pdf"
    list2.append(newentry)
 
for item in source_dir:
    if item.endswith('pdf'):
        name = item.replace('.pdf','')
        name2 = int(name)
        list.append(name2)
        
for newentry in list2:
    merger.append(newentry)
    
print(list2)
        
print("Completed Merge")

 
merger.write('Completed Merge.pdf')       
merger.close()

