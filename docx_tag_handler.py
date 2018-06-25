######################################################
#
# docx_tag_handler - replaces tagged docx text with substitutes
# # written by Jonas Malmström (jonas@mit-automation.se)
#
######################################################

import re
import os
import contextlib
from zipfile import ZipFile


class DocxTagHandler:

    def __init__(self, tag_start, tag_end):
        self.regex_pattern = re.escape(tag_start) + r'[a-zA-Z0-9]+' + re.escape(tag_end)

    def get_tags(self, src_file):

        with ZipFile(src_file).open('word/document.xml') as document:
            text = document.read()
            return re.findall(self.regex_pattern, text)


    def replace_tags(self, src_file, target_file, sub):

        with contextlib.suppress(FileNotFoundError):
            os.remove(target_file)

        input_zipfile = ZipFile(src_file)
        outputZipFile = ZipFile(target_file, "a")
        macros = {'(%pris%)': '500', '(namn)': 'Jonas'}

        with input_zipfile.open('word/document.xml') as zippedXMLText:
            XMLtext = str(zippedXMLText.read())
            XMLtext = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + XMLtext[61:-1]
        for item, key in macros.items():
            XMLtext = re.sub(key, item, XMLtext)

        for file in input_zipfile.filelist:
            if file.filename == "word/document.xml":
                outputZipFile.writestr(file.filename, XMLtext)

            else:
                outputZipFile.writestr(file.filename, input_zipfile.read(file))

        input_zipfile.close()
        outputZipFile.close()

        ZipFile('testin.docx').extractall('./original')
        ZipFile('testout.docx').extractall('./modified')
