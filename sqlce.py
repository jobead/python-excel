import os
import time
import datetime

import zipfile

import xml.dom.minidom 

def get_files_from_folder(path = os.environ['HOME']):
	homedir = os.environ['HOME'] + '/python'

def get_external_data_connections(path = os.environ['HOME'] + '/python/excelfiles', filename = 'external_data_connections'):
	
	print 'STARTING SEARCH FOR EXTERNAL DATA CONNECTIONS'

	homedir = os.environ['HOME'] + '/python'

	print 'DIRECTORY TO SEARCH'
	print '\t' + path
	os.chdir(path)

	filelist = os.listdir(path)
	for files in filelist:
		i=0
		if filelist[i] == '.DS_Store':  # eliminates apple's directory view settings file from list
			filelist[i:i+1] = []
		i=i+1
	
	#print filelist

	print 'FILES TO SEARCH'
	for files in filelist:
		print '\t' + files 

	outputfile = open(homedir + '/' + filename, 'w+')

	print outputfile

	outputfile.write('THESE ARE ALL MY EXCEL FILES\t\t\tRUNTIME: ' + str(datetime.datetime.now()) + '\n\n')

	print 'i am about to start writing files'

	i=0
	for files in filelist:
		outputfile.write('-- ' + filelist[i]+'\n\n')


		archive = zipfile.ZipFile(path + '/' + filelist[i])
		conn = archive.read(name='xl/connections.xml')
		
		# basicxml = xml.dom.minidom.parseString(books)
		# xmlelements = basicxml.getElementsByTagName('sheet')
		
		# sheetrowsout = list(len(xmlelements)*' ')


		basicxml = xml.dom.minidom.parseString(conn) 
		# conn = str(basicxml.toprettyxml())
		xmlelements = basicxml.getElementsByTagName('connection')

		# print 'i have parsed my xml'

		connrowsout = list(len(xmlelements)*' ')

		# print 'i have setup my output list'

		# print xmlelements

		j=0
		for conns in xmlelements:
			# print 'i am about to parse my ' + str(j) + 'th xmlelement'
			# print xmlelements[j].attributes['id'].value + '\t\t' + xmlelements[j].attributes['name'].value

			conncommand = conns.getElementsByTagName('dbPr')
			querytext = '\t\t' + str(conncommand[0].attributes['command'].value)
			querytext = str.replace(querytext,'_x000d__x000a_','\n\t\t')
			querytext = str.replace(querytext,'_x0009_','\t')
			querytext = str.replace(querytext,'&gt;','>')
			querytext = str.replace(querytext,'&lt;','<')

			# print conncommand
			
			connrowsout[j] = '--' + '\t' + xmlelements[j].attributes['name'].value + '\n' + querytext + '\n\n'
			# print 'i have parsed my ' + str(j) + 'th element'
			j=j+1
		
		connrowsout.sort()
		#print pretty_xml_as_string
		
		



		# conn = str.replace(conn,'command="','command="\n')
		
		k=0
		for outputs in connrowsout:
			outputfile.write(connrowsout[k]+'\n')
			k=k+1

		outputfile.write('\n\n\n')	
		i=i+1


	outputfile.write('\n\n\n')
	

	print 'i have written all my files'

	#print os.getcwd()

	#archive.printdir()

	
	
	#print conn1[:100] + '... plus ' + str(len(conn1) - 100) + ' more characters'
	
	os.chdir(homedir)

	print 'i have reset the working directory to ' + homedir

	
	#print os.getcwd()




def get_sheet_names(path = os.environ['HOME'] + '/python/excelfiles', filename = 'sheet_names'):
	
	print 'STARTING SEARCH FOR SHEET NAMES'

	homedir = os.environ['HOME'] + '/python'

	print 'DIRECTORY TO SEARCH'
	print '\t' + path
	os.chdir(path)

	filelist = os.listdir(path)
	for files in filelist:
		i=0
		if filelist[i] == '.DS_Store':  # eliminates apple's directory view settings file from list
			filelist[i:i+1] = []
		i=i+1
	
	#print filelist

	print 'FILES TO SEARCH'
	for files in filelist:
		print '\t' + files 

	outputfile = open(homedir + '/' + filename, 'w+')

	print outputfile

	outputfile.write('THESE ARE ALL MY EXCEL FILES\n\n')

	print 'i am about to start writing files'

	i=0
	for files in filelist:
		print filelist[i]

		outputfile.write(filelist[i]+'\n')

		archive = zipfile.ZipFile(path + '/' + filelist[i])

		# print archive.infolist()
		# print archive.namelist()

		books = archive.read(name='xl/workbook.xml')

		
		basicxml = xml.dom.minidom.parseString(books)
		xmlelements = basicxml.getElementsByTagName('sheet')
		
		sheetrowsout = list(len(xmlelements)*' ')

		j=0
		for sheetrows in xmlelements: 
			# print (try: visible = xmlelements[j].attributes['state'].value except: 'hidden') + '\t' + xmlelements[j].attributes['name'].value
			try:
				visible = '\t\t' + xmlelements[j].attributes['state'].value
			except:
				visible = '' 
				# print '\t\t\t' + 'hidden'

			sheetrowsout[j] = visible.strip()[:1] + ('000' + xmlelements[j].attributes['r:id'].value[3:])[-4:] + '\t' + visible + '\t' + xmlelements[j].attributes['name'].value + '\n'


			j=j+1

		sheetrowsout.sort()


		# books = str(basicxml.toprettyxml())

		k=0
		for outputs in sheetrowsout:
			outputfile.write(sheetrowsout[k])
			k=k+1

		outputfile.write('\n\n')

		i=i+1



	os.chdir(homedir)

	print 'i have reset the working directory to ' + homedir

	
















