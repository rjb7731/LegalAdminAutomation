import os
import PyPDF2
import numpy as np

listtext = []
docs = {}

number = len(os.listdir(r'FILE LINK HERE'))

def pdftoarray():
        directory = r'FILE LINK HRE'
        for entry in os.scandir(directory):
                print(entry)
                if (entry.path.endswith(".pdf")):    
                    pdfFileObj = open(entry.path, 'rb')
                    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                    count = pdfReader.numPages
                    output = []
                    for i in range(count):
                            page = pdfReader.getPage(i)
                            pagereturn = page.extractText()
                            textamend = pagereturn.lower().replace("\n", "")
                            output.append(textamend)
                            #print("Document page count:"+str(count))
                    newoutput = ' '.join(output)
                    listtext.append(newoutput) 

        for x in range(len(listtext)):
                output = {x:listtext[x]}
                docs.update(output)

        print('Completed')
                                                    
                
pdftoarray()
              
