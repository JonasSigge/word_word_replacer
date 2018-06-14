"""main file, acting as central controller with external functions
"""


import re
import os
import contextlib
from zipfile import ZipFile

with contextlib.suppress(FileNotFoundError):
    os.remove('./testout.docx')

inputZipFile = ZipFile('testin.docx')
outputZipFile = ZipFile('testout.docx',"a")
macros = {'(%pris%)':'500','(namn)':'Jonas'}

with inputZipFile.open('word/document.xml') as zippedXMLText:
    XMLtext = str(zippedXMLText.read())
    XMLtext = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + XMLtext[61:-1]
for item, key in macros.items():
    XMLtext = re.sub(key,item,XMLtext)


for file in inputZipFile.filelist:
    if file.filename == "word/document.xml":
        outputZipFile.writestr(file.filename, XMLtext)

    else:
        outputZipFile.writestr(file.filename, inputZipFile.read(file))


inputZipFile.close()
outputZipFile.close()

ZipFile('testin.docx').extractall('./original')
ZipFile('testout.docx').extractall('./modified')
