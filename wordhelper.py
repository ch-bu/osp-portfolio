#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your
work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
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

def get_data(title, iter_paragraphs):


	if title == 'Zu Ihrer Person':
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