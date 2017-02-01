#!/usr/bin/env python
"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your
work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

import sys
import zipfile
import re
# http://stackoverflow.com/questions/41550620/python-docx-get-info-from-dropdownlist-in-table

from docx import Document
from bs4 import BeautifulSoup

if __name__ == '__main__':
    try:
        document = Document(sys.argv[1])
        newfile = open(sys.argv[2], 'w')
    except:
        print(
            "Please supply an input and output file. For example:\n"
            "  example-extracttext.py 'My Office 2007 document.docx' 'outp"
            "utfile.txt'"
        )
        exit()

    # Fetch all the text out of the document we just created
    # Make explicit unicode version
    newparatextlist = []
    for paratext in document.paragraphs:
        newparatextlist.append(paratext.text)

    # Print out text of document with two newlines under each paragraph
    newfile.write('\n\n'.join(str(v) for v in newparatextlist))

    # Get dropdown
    dropdown = zipfile.ZipFile(sys.argv[1])
    xml_data = dropdown.read('word/document.xml')
    dropdown.close()

    soup = BeautifulSoup(xml_data, 'xml')
    dropdownList = soup.findAll('sdtContent')

    sdtContentElements = []
    for element in dropdownList:
        sdtContentElements.append(element)
    
    #activity = re.sub('<[^>]*>', '', str(element.findAll('t')[0]))
    #print(activity)

