#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import zipfile
from bs4 import BeautifulSoup

"""
Put description here
"""

def get_simple_data(paragraphs):
    # Init result string
    result = ''

    # Loop over every paragraph
    for paragraph in paragraphs:
        # Append new paragraph to previous paragraph
        result = result + ' ' + paragraph.text

    # Return string of paragraphs
    return result.replace('\n', ' ').replace('\r', '')


def get_person(person_string):
    # Gets the values of the specific person
    first_name = re.search('Vorname:(.+?)Nachname', person_string).group(1).strip()
    last_name = re.search('Nachname:(.+?)Matrikelnummer', person_string).group(1).strip()
    number = re.search('Matrikelnummer:(.+?)E-Mail', person_string).group(1).strip()
    mail = re.search('E-Mail:(.+?)Hauptfach', person_string).group(1).strip()
    subject_one = re.search('Hauptfach 1:(.+?)Hauptfach', person_string).group(1).strip()
    subject_two = re.search('Hauptfach 2:(.+?)$', person_string).group(1).strip()

    result = {'Vorname': first_name, 'Nachname': last_name, 'Matrikelnummer': number, \
        'Mail': mail, 'Hauptfach_1': subject_one, 'Hauptfach_2': subject_two}

    return result

def get_activity(document, iter_paragraphs, word_file):
    # Get dropdown
    zip_file = zipfile.ZipFile(word_file)
    xml_data = zip_file.read('word/document.xml')
    zip_file.close()

    soup = BeautifulSoup(xml_data, 'xml')
    dropdownList = soup.findAll('sdtContent')

    sdtContentElements = []
    for element in dropdownList:
        sdtContentElements.append(element)

    try:
        activity = re.sub('<[^>]*>', '', str(element.findAll('t')[0]))
    except UnboundLocalError:
        activity = ''

    return activity

def get_data(title, iter_paragraphs, document, word_file):
    # Check which word file is analyzed and behave
    # appropriately
    if title == 'Zu Ihrer Person':
        return get_person(get_simple_data(iter_paragraphs))
    elif title == 'Beobachten. Erste Tätigkeit' or \
        title == 'Beobachten. Zweite Tätigkeit' or \
        title == 'Beobachten. Dritte Tätigkeit':
        activity = get_activity(document, iter_paragraphs, word_file)
        content = get_simple_data(iter_paragraphs)
        return {'activity': activity, 'content': content}
    elif title == 'Begleitung Alltag Lehrperson' or \
        title == 'Wahlbeobachtung' or \
        title == 'Erste Durchführung einer zentralen Tätigkeit' or \
        title == 'Zweite Durchführung einer zentralen Tätigkeit' or \
        title == 'Wahl-Aufgabe' or \
        title == 'Interview mit einer Lehrkraft' or \
        title == 'Schlüsselsituation':
        print(title)
        return get_simple_data(iter_paragraphs)
    else:
        return None
