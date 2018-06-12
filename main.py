"""main file, acting as central controller with external functions
"""


import re

from zipfile import ZipFile


inputZipFile = ZipFile('testin.docx')
outputZipFile = ZipFile('testout.docx',"a")
macros = {'(%pris%)':'500','(%namn%)':'Jonas'}

with inputZipFile.open('word/document.xml') as zippedXMLText:
    XMLtext = str(zippedXMLText.read())
        
for item, key in macros.items():
    XMLtext = re.sub(key,item,XMLtext)


for file in inputZipFile.filelist:
    if not file.filename == "word/document.xml":
        outputZipFile.writestr(file.filename, XMLtext)
    else:
        outputZipFile.writestr(file.filename, inputZipFile.read(file))
    

inputZipFile.close()
outputZipFile.close()


