#!/usr/bin/env python
"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your
work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

# http://stackoverflow.com/questions/41550620/python-docx-get-info-from-dropdownlist-in-table

import sys
import zipfile
import re
import glob
import os
import shutil
from docx import Document
from bs4 import BeautifulSoup


def unzip_files(compressed_file):
    # Unzip file
    zip_ref = zipfile.ZipFile(compressed_file, 'r')
    zip_ref.extractall(temp_path)
    zip_ref.close()

    return None

def analyze_files(temp_path):

    word_files = [word_file for word_file in glob.iglob(temp_path + '\\**\\*.docx', recursive=True)]

    print(word_files)

# Run code only if the program is run by itself and not
# when it is imported from another module
if __name__ == '__main__':
    
    # Get dir of portfolio folder two
    file_path = os.path.dirname(os.path.realpath(__file__))
    portfolio_path = file_path + '\\portfolios-two'
    temp_path = file_path + '\\temp'

    # Get all compressed files and save path in zip_files
    file_extensions = ['.zip', '.tar', '.7s']
    zip_files = [portfolio for portfolio in os.listdir(portfolio_path) if portfolio.endswith(tuple(file_extensions))]

    # Unzip all files
    unzipped_files = list(map(lambda compressed_file: unzip_files(portfolio_path + '\\' + compressed_file), zip_files))

    # Analyze all files
    results = analyze_files(temp_path)

    # try:
    #     document = Document(sys.argv[1])
    #     newfile = open(sys.argv[2], 'w')
    # except:
    #     print(
    #         "Please supply an input and output file. For example:\n"
    #         "  example-extracttext.py 'My Office 2007 document.docx' 'outp"
    #         "utfile.txt'"
    #     )
    #     exit()

    # Fetch all the text out of the document we just created
    # Make explicit unicode version
    # newparatextlist = []
    # for paratext in document.paragraphs:
    #     newparatextlist.append(paratext.text)

    # # Print out text of document with two newlines under each paragraph
    # newfile.write('\n\n'.join(str(v) for v in newparatextlist))

    # # Get dropdown
    # dropdown = zipfile.ZipFile(sys.argv[1])
    # xml_data = dropdown.read('word/document.xml')
    # dropdown.close()

    # soup = BeautifulSoup(xml_data, 'xml')
    # dropdownList = soup.findAll('sdtContent')

    # sdtContentElements = []
    # for element in dropdownList:
    #     sdtContentElements.append(element)
    
    #activity = re.sub('<[^>]*>', '', str(element.findAll('t')[0]))
    #print(activity)


