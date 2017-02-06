#!/usr/bin/env python
# -*- coding: utf-8 -*-

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
from wordhelper import get_data
import json
import csv


def unzip_files(compressed_file):
    # Unzip file
    zip_ref = zipfile.ZipFile(compressed_file, 'r')
    zip_ref.extractall(temp_path)

    # Analyze data
    subject_results = analyze_files(temp_path)

    # Delete directory for new subject
    shutil.rmtree(temp_path)

    # Close zip file
    zip_ref.close()

    # Return results
    return subject_results

def return_file_content(document, word_file):

    # Store every result in dictionary that
    # can be processed later
    results = {}

    # Add title of document to results
    title = document.paragraphs[0].text
    results[title] = ''

    # Make sure not to include the first paragraph
    # that includes the title
    iter_paragraphs = iter(document.paragraphs)
    next(iter_paragraphs)

    # Get data from current word file
    results[title] = get_data(title, iter_paragraphs, document, word_file)

    # Return content of word file in dictionary
    return results

def analyze_files(temp_path):

    # Get all docx files
    word_files = [word_file for word_file in glob.iglob(temp_path + '\\**\\*.docx', recursive=True)]

    # Store results of all word files in this variable
    res = []

    # Loop over every word file
    for word_file in word_files:
        # Get word document
        document = Document(word_file)

        res.append(return_file_content(document, word_file))

    return res


# Run code only if the program is run by itself and not
# when it is imported from another module
if __name__ == '__main__':
    
    # Get dir of portfolio folder two
    file_path = os.path.dirname(os.path.realpath(__file__))
    portfolio_path = file_path + '\\portfolios-two'
    temp_path = file_path + '\\temp'

    # Get all compressed files and save path in zip_files
    # file_extensions = ['.zip', '.tar', '.7s']
    file_extensions = ['.zip']
    zip_files = [portfolio for portfolio in os.listdir(portfolio_path) if portfolio.endswith(tuple(file_extensions))]

    # Get data from all subjects
    data_subjects = list(map(lambda compressed_file: unzip_files(portfolio_path + '\\' + compressed_file), zip_files))

    # print(data_subjects)

    # json.dump(data_subjects, open('test.txt', 'w'))


    # Key keys
    keys = [list(word_file.keys())[0] for word_file in data_subjects[0]]
    print(keys)

    # # Write results to disk
    # with open('mycsvfile.csv', 'wb') as f:  # Just use 'w' mode in 3.x

    #     for subject in data_subjects:
    #         # print(subject)

    #         for word_file in subject:
    #             print(word_file.keys())
    #             # print(subject[1].keys())
    #             # w = csv.DictWriter(f, word_file.keys())
    #             # w.writeheader()
    #             # w.writerow(word_file)

