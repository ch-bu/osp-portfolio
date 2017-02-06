#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re

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
	return result


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

def get_data(title, iter_paragraphs, document):
	# Check which word file is analyzed and behave
	# appropriately
	if title == 'Zu Ihrer Person':
		print(get_person(get_simple_data(iter_paragraphs)))
		return 'Ich bin\'s halt'
	elif title == 'Beobachten. Erste Tätigkeit':
		return 'erste Tätigkeit'
	elif title == 'Beobachten. Zweite Tätigkeit':
		return 'zweite Tätigkeit'
	elif title == 'Beobachten. Dritte Tätigkeit':
		return 'dritte Tätigkeit'
	elif title == 'Begleitung Alltag Lehrperson':
		return 'Begleitung Alltag'
	elif title == 'Wahlbeobachtung' or \
		 title == 'Erste Durchführung einer zentralen Tätigkeit' or \
		 title == 'Zweite Durchführung einer zentralen Tätigkeit' or \
		 title == 'Wahl-Aufgabe' or \
		 title == 'Interview mit einer Lehrkraft' or \
		 title == 'Schlüsselsituation':
		return get_simple_data(iter_paragraphs)
	else:
		return None


