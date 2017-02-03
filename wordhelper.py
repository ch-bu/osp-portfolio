#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your
work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

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
	elif title == 'Wahlbeobachtung':
		return 'Wahlbeobachtung'
	elif title == 'Erste Durchführung einer zentralen Tätigkeit':
		return 'Erste Durchführung zentrale Tätigkeit'
	elif title == 'Zweite Durchführung einer zentralen Tätigkeit':
		return 'Zweite Durchführung zentrale Tätigkeit'
	elif title == 'Wahl-Aufgabe':
		return 'Wahl-Aufgabe'
	elif title == 'Interview mit einer Lehrkraft':
		return 'Interview Lehrkraft'
	elif title == 'Schlüsselsituation':
		return 'Schlüsselsituation'
	else:
		return None