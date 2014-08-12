#!/usr/bin/python

"""
# RTF2DOCX.PY
#
#   Iyad Obeid, 8/7/2014, v1.0.1
#
# Converts rtf to docx
#   Captures text in an rtf file including headers and footers, and converts to a docx
#   Run with -h or -help flag for more information on how to run
#
#   Code is based on docx.py which is downloaded from here:
#       https://github.com/mikemaccana/python-docx
#   and pyth 0.6.0 which is downloaded from here:
#       https://pypi.python.org/pypi/pyth/
#
#   Installation requires
#     apt-get install libxml2-dev libxslt1-dev python-dev
#       (this may be required on linux, shouldn't be necessary on later
#        model osx systems)
#
#     sudo pip install lxml
#     sudo pip install Pillow (used to be PIL)
"""

from rtf_parser.reader import Rtf15Reader
from rtf_parser.docx import *

import sys
import os


def main():
	# main function

	# parse the command line, check for errors
	flag, fileNameInput, fileNameOutput = init_vars()

	# print the help screen if requested
	if flag['help'] is True:
		print_help_screen()
		exit()

	# print the verbose screen if requested
	if flag['verbose'] is True:
		print_verbose_screen(fileNameInput, fileNameOutput)

	# read the rtf file
	doc = read_the_rtf_file(fileNameInput)

	# take the extracted text and create a docx
	create_docx(doc, fileNameOutput)


def init_vars():

	flag = dict(verbose=False, help=False)
	theArgs = sys.argv[1:]
	nArguments = len(sys.argv)-1

	# check all the input switches in order to set up process flow properly
	for i in range(1, len(sys.argv)):

		if (sys.argv[i].lower() == '-verbose') or \
			(sys.argv[i].lower() == '-v'):
			flag['verbose'] = True
			nArguments -= 1
			theArgs.pop(0)

		elif (sys.argv[i].lower() == '-help') or \
			(sys.argv[i].lower() == '-h'):
			flag['help'] = True
			nArguments -= 1
			theArgs.pop(0)

		# unknown switch
		elif sys.argv[i][0] == '-':
			print(' ')
			print('ERROR: switch ' + sys.argv[i].upper() + ' not found')
			print('	Try ''./rtf2docx.py -help'' for more options')
			print(' ')
			exit()

	# Check to see if the minimum number of arguments (2) has been
	# supplied. Note that you don't need two arguments if the help
	# flag has been thrown
	if (flag['help'] is False) and (nArguments == 0):
		print(' ')
		print('ERROR: provide at least an input filename')
		print(' ')
		exit()

	# extract the filenames
	fileNameInput = theArgs[0]
	if nArguments == 1:
		fileNameOutput=fileNameInput[0:-3]+'docx'
	else:
		fileNameOutput = theArgs[1]

	# check to see if the specified input file exists
	if os.path.isfile(fileNameInput) is False:
		print (' ')
		print ('ERROR: input file ' + fileNameInput + ' not found')
		print (' ')
		exit()

	return flag, fileNameInput, fileNameOutput


def print_help_screen():
	print(' ')
	print('RTF2DOCX.py : coverts an RTF file to an MSWord docx file')
	print('    ./rtf2docx.py inputfile.docx outputfile.txt')
	print('    optional switches: -verbose (-v), -help')
	print(' ')


def print_verbose_screen(fileNameInput, fileNameOutput):
		print 'rtf2docx:'
		print '  Input file:  ', fileNameInput
		print '  Output file: ', fileNameOutput


def read_the_rtf_file(fileNameInput):
	return Rtf15Reader.read(open(fileNameInput, "rb"))


def create_docx(doc, fileNameOutput):

	# this part sets up the docx
	relationships = relationshiplist()
	document = newdocument()
	body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

	# helpers to deal with special cases
	specialText = []
	specialFlag = False
	firstFlag = True

	# iterate through every element that has been detected in the rtf text
	# the basic case is that its just a simple paragraph and it gets outputted
	# however we have to deal with a couple special cases
	# bold and italic text were being treated as their own paragraph
	# this code takes bold/italic, stores it temporarily, and then appends the following
	# paragraph. That way, you don't get a new paragraph immediately after every bold/ital
	# we also have to catch 'cell' to deal with table and force a print every time.
	for element in doc.content:
		for item in element.content:
			if 'jpegblip' in item.properties:
				f = open('foo.jpg','wb')
				f.write(item.content[0].decode('hex'))
				f.close()
				relationships,picpara=picture(relationships,'foo.jpg','some text')
				body.append(picpara)
				# os.remove('foo.jpg')
			elif 'cell' in item.properties:
				if specialText:
					body.append(paragraph(specialText))
					specialText = []
				body.append(paragraph(item.content))
			elif ('bold' in item.properties) or ('italic' in item.properties):
				specialText.extend(item.content)
				specialFlag = True
			else:
				if specialFlag is True:
					specialText.extend(item.content)
					body.append(paragraph(specialText))
					specialText = []
					specialFlag = False
				else:
					if firstFlag and item.content[0] == u'':
						''
					else:
						body.append(paragraph(item.content))
			firstFlag = False

	# Create our properties, contenttypes, and other support files
	title	= 'Python docx demo'
	subject  = 'A practical example of making docx from Python'
	creator  = 'Mike MacCana'
	keywords = ['python', 'Office Open XML', 'Word']

	_coreprops = coreproperties(title=title, subject=subject, creator=creator,
								keywords=keywords)
	_appprops = appproperties()
	_contenttypes = contenttypes()
	_websettings = websettings()
	_wordrelationships = wordrelationships(relationships)

	# Save our document
	savedocx(document, _coreprops, _appprops, _contenttypes, _websettings,
			_wordrelationships, fileNameOutput)

if __name__ == '__main__':
	main()
