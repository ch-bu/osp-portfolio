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


    # Write results to disk
    with open('mycsvfile.csv', 'w') as f:

        # Create writer for csv file
        writer = csv.writer(f, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')

        # Write header
        writer.writerow(['Vorname', 'Nachname', 'Matrikelnummer',
                        'Mail', 'Hauptfach.1', 'Hauptfach.2',
                        'Begleitung Alltag Lehrperson', 'beobachten.one.activity',
                        'beobachten.one.content', 'beobachten.two.activity',
                        'beobachten.two.content', 'beobachten.three.activity',
                        'beobachten.three.activity', 'beobachten.three.content',
                        'Wahlbeobachtung',
                        'Erste Durchführung einer zentralen Tätigkeit',
                        'Zweite Durchführung einer zentralen Tätigkeit',
                        'Schlüsselsituation',
                        'Interview mit einer Lehrkraft',
                        'Wahl-Aufgabe'])

        # Loop over every subject
        for subject in data_subjects:

            my_dict = {}

            # Loop over every file
            for word_file in subject:
                # Get key of file
                key = list(word_file.keys())[0]

                # Add personal information
                if key == 'Zu Ihrer Person':
                    curr_dict = word_file[key]
                    my_dict['Vorname'] = curr_dict['Vorname']
                    my_dict['Nachname'] = curr_dict['Nachname']
                    my_dict['Matrikelnummer'] = curr_dict['Matrikelnummer']
                    my_dict['Mail'] = curr_dict['Mail']
                    my_dict['Hauptfach.1'] = curr_dict['Hauptfach_1']
                    my_dict['Hauptfach.2'] = curr_dict['Hauptfach_2']
                    # my_dict = {**word_file[key]}
                if key == 'Beobachten. Dritte Tätigkeit':
                    curr_dict = word_file[key]
                    my_dict['beobachten.three.activity'] = curr_dict['activity']
                    my_dict['beobachten.three.content'] = curr_dict['content']
                elif key == 'Beobachten. Zweite Tätigkeit':
                    curr_dict = word_file[key]
                    my_dict['beobachten.two.activity'] = curr_dict['activity']
                    my_dict['beobachten.two.content'] = curr_dict['content']
                elif key == 'Beobachten. Erste Tätigkeit':
                    curr_dict = word_file[key]
                    my_dict['beobachten.one.activity'] = curr_dict['activity']
                    my_dict['beobachten.one.content'] = curr_dict['content']
                else:
                    my_dict[key] = word_file[key]

            # Write row for subject
            writer.writerow([my_dict['Vorname'], my_dict['Nachname'], my_dict['Matrikelnummer'],
                            my_dict['Mail'], my_dict['Hauptfach.1'], my_dict['Hauptfach.2'],
                            my_dict['Begleitung Alltag Lehrperson'], my_dict['beobachten.one.activity'],
                            my_dict['beobachten.one.content'], my_dict['beobachten.two.activity'],
                            my_dict['beobachten.two.content'], my_dict['beobachten.three.activity'],
                            my_dict['beobachten.three.activity'], my_dict['beobachten.three.content'],
                            my_dict['Wahlbeobachtung'],
                            my_dict['Erste Durchführung einer zentralen Tätigkeit'],
                            my_dict['Zweite Durchführung einer zentralen Tätigkeit'],
                            my_dict['Schlüsselsituation'],
                            my_dict['Interview mit einer Lehrkraft'],
                            my_dict['Wahl-Aufgabe']])
            # w = csv.DictWriter(f, my_dict.keys(), delimiter=';', lineterminator='\n')
            # w.writerow(my_dict)

        # w.writeheader()
