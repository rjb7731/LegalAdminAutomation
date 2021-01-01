import re
import pickle
import os
import PyPDF2

doclist = []
entitylist = []

f = open(r'\Documentlist.txt', 'r')
docnames = f.readlines()
for x in docnames:
    docname = x.strip().lower()
    doclist.append(docname)

f2 = open(r'\entitylist.txt', 'r')
entitynames = f2.readlines()
for x in entitynames:
    entityname = x.strip().lower()
    entitylist.append(entityname)

listtext = []
docsdic = {}
docfile = r"FOLDER OF PDF'S GOES HERE"
number = len(os.listdir(docfile))
filetypes = ['.pdf','.PDF']

filelist =[]

def pdftoarray():
        directory = r"FOLDER OF PDF'S GOES HERE"
        for entry in os.scandir(directory):
                if entry.path.endswith(tuple(filetypes)):
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
                    filelist.append(entry)
                else:
                    print(f"-FAIL - Cannot read all PDF - {entry}")

        for x in range(len(listtext)):
                output = {x:listtext[x]}
                docsdic.update(output)

        print('****Completed - Doc to array process***')                             

    except:
        return "*NO ENTITY FOUND - TRY AGAIN"

def findentity(string):
    """Searches entity names from list"""
    for x in entitylist:
        if x in string:
            print(f"(Doc.{i})--Entity = {x.title()}")
            break

def finddocname(string):
    """Searches doc names from list"""
    for x in doclist:
        foundvar = f"-->Doc name = {x.title()}"
        if x in string:
            print(foundvar)
            break

pdftoarray()

for i in range(number):
    entry = docsdic[i]
    entrysnip = entry[0:750]
    print("-------------------")
    print(f"File: {filelist[i]}")
    finddocname(entrysnip)
