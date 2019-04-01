#!/usr/bin/python3
# -*- coding: utf-8 -*-
""" This python script will read in an Excel sheet which is the output of eolNetworkSummary.py - my
	February challenge submission which analyzed an NMSAAS output to list the software versions in use on a network.
	It will then gather the vulnerablities associated with each software version and add this information
	to the Excel sheet provided as the input.

	Additionally - there needs to be a file in the same directory as the script named auth.txt where first line is the 
	Cisco provided Client ID and the second line is the Cisco provided Client Secret to authenticate and get tokens.
	
	Example of auth.txt:
		sfjlsfjsdklfjasdkljfsladk
		askldfjlaksfjlskdfjklasdf
		
	In this case "sfjlsfjsdklfjasdkljfsladk" would be the Cisco provided 'Client ID' and "askldfjlaksfjlskdfjklasdf" would
	be the provided 'Client Secret'.

	Arguments:
		1) Filename (ends in .xls) - this is an output file from my February submission eolNetworkSummary.py.
	
	The output information is added to the Excel sheet in additional tabs.
		
"""

# import modules HERE
import sys											# this allows us to analyze the arguments	
import os											# this allows us to check on the file
import xlrd											# this allows us to import an Excel file
import xlwt											# this allows us to output data to an Excel file
from xlutils.copy import copy as excel_copy_rdwt	# this allows a workbook read in to be converted to a workbook that can be written
from datetime import datetime						# useful for getting timing information and for some data translation from Excel files
from contextlib import contextmanager
import requests										# for web queries and API calls
import json											# for translating data back and forth from web requests

# additional information about the script
__filename__ = "getCiscoSOftwareVul.py"
__author__ = "Robert Hallinan"
__email__ = "rhallinan@netcraftsmen.com"

#
# version history
#


"""
	20190207 - Initially creating the script
"""

@contextmanager
def open_file(path, mode):
	the_file = open(path, mode)
	yield the_file
	the_file.close()

def importFile(passedArgs):
	""" this script will determine which function to run when parsing the input file, import the data, and return a list of dictionaries
	"""

 	# assign variables
	fileInput=str(passedArgs)

	# Does the file exist?
	if not os.path.exists(fileInput):
		print("File name provided to convert does not exist. Closing now...")
		sys.exit()
		
	if fileInput[-4:].lower() == ".xls" or fileInput[-5:].lower() == ".xlsx":
		print("File Input is: "+fileInput)
		return parseExcel(fileInput)
	else:
		sys.exit()
	

def parseExcel(fileInput):
	""" This function parses the Excel file and returns a list of dictionaries with the keys of each dictionary as the column header and the value specific to the row
	"""	
	print("Network information will be parsed from the Excel input file....")

	#
	# define outputs
	#

	# make a list for the items in the file
	outputNetDev=[]
	
	# open the Excel file
	newBook = xlrd.open_workbook(fileInput)
	
	# get a list of the sheet names, make sure that the sheet named 'sort-by-count' or 'sort-by-eos' is present
	sheetNames=[]
	for sheet in newBook.sheets():
		sheetNames.append(sheet.name)
	if 'sort-by-count' not in sheetNames and 'sort-by-eos' not in sheetNames:
		print("This Excel sheet is not valid. Exiting...")
		sys.exit()
	
	if 'sort-by-count' in sheetNames:
		sheetInput = 'sort-by-count'
	else:
		sheetInput = 'sort-by-eos'
	
	# get the correct sheet out of the workbook
	sheetParse = newBook.sheet_by_name(sheetInput)

	# declare some general info so it is accessible for multiple iterations of the for loop once initially modified
	colHeaderList=[]
	colHeaderRead=sheetParse.row_values(0)
	for f in range(len(colHeaderRead)):
		colHeaderList.append((f,colHeaderRead[f].rstrip()))	

	for newDevRow in range(2,sheetParse.nrows):
		# declare empty dictionary that we can add to for this item's information
		newItem={}
		
		# get a list of the line since is CSV
		itemList=sheetParse.row_values(newDevRow)
		# print(itemList)
		
		# check the line - if the length of fields is longer than the length of columns then there is a comma somewhere in the entry
		# user has to make sure that there is no comma
		if len(itemList) != len(colHeaderList):
			print("One or more items have a comma in their value string which makes this impossible to properly parse as a CSV.")
			sys.exit()
		
		# get the info on this item
		for pair in colHeaderList:
			
			# does this need to be a datestamp translated?
			if 'Created' in pair[1] or 'End-of' in pair[1]:
				# print("testing on row: "+str(newDevRow))
				# print("value testing: "+str(itemList[pair[0]]))
				try:
					newItem[pair[1]]=str(datetime(*xlrd.xldate_as_tuple(itemList[pair[0]], newBook.datemode))).rstrip()
				except:
					# print("this didn't work")
					# must be an empty field
					newItem[pair[1]]=''
			else:
				newItem[pair[1]]=str(itemList[pair[0]]).rstrip()
		# assign the dictionary of the new item to the list
		outputNetDev.append(newItem)

	return outputNetDev
	
def outputExcel(listOutput,fileName,tabName):
	""" listOutput: this should be a list of lists; first item should be header file which should be written.
		fileName: Name of the Excel file to which this data should be written
		tabName: Name of the tab to which this data should be written
	"""
	
	# since before this would get called - it is assumed that the file was initialized - if the file now exists it is because another
	# tab is already in it from this script - thus check to see if the file is there - if so then just open workbook using xlrd
	if os.path.exists(fileName):
		outBook = xlrd.open_workbook(fileName)
		outBookNew = excel_copy_rdwt(outBook)
		outBook = outBookNew
	else:	
		# make the new Workbook object
		outBook = xlwt.Workbook()

	# add the sheet with the tab name specified
	thisSheet = outBook.add_sheet(tabName)
	
	# get number of columns
	numCols=len(listOutput[0])
	
	for rowNum in range(len(listOutput)):
		writeRow = thisSheet.row(rowNum)
		# print(listOutput[rowNum])
		for x in range(numCols):
			writeRow.write(x,str(listOutput[rowNum][x]))
			
	# save it to the Excel sheet at the end
	outBook.save(fileName)

def getAuthToken(uname, upassword):
	# this function will return a new authentication token based on provided username and password
	# and as well the new headers to use for the request
	
	# extract the authentication information
	try:
		with open_file('auth.txt','r') as fileIn:
			[ uname, upassword ] = [ x.rstrip() for x in fileIn.readlines() ]
	except:
		print("Required auth.txt file with info does not exist. Exiting...")
		sys.exit()
	
	# setup the auth token request
	headersAuth = {
				'Content-Type': 'application/x-www-form-urlencoded'
			  }
	dataPayloadAuth = {
					'client_id': uname,
					'client_secret': upassword,
					'grant_type': 'client_credentials'
				  }
	authenURL = "https://cloudsso.cisco.com/as/token.oauth2"
	
	# do the initial request to get a token
	response = requests.post(authenURL, headers = headersAuth, data = dataPayloadAuth)	

	if response.status_code == 200:
		# extract the access token
		accessToken = json.loads(response.text)['access_token']
	else:
		print("No access token. Exiting...")
		sys.exit()

	requestHeader = {
				'Accept': 'application/json',
				'Authorization': 'Bearer ' + accessToken,
			  }	
		
	return accessToken, requestHeader
	
def main(system_arguments):

	# get a python list of dictionaries by parsing the CSV file - validate that there is even an argument there using try
	try:
		networkInventory = importFile(system_arguments[1])
	except:
		print("No valid argument of a filename provided. Exiting...")
		sys.exit()
	
	# set the output filename
	outputExcelFile = system_arguments[1]

	# get a list of the software items
	softwareList = [ x['Software'] for x in networkInventory ]
	# print(softwareList)
	
	# let user know that this whole effort will only be run for devices running IOS or IOS XE - Nexus is not supported
	print("The vulnerabilities will now be gathered. The Cisco API only supports IOS and IOS XE - Nexus devices are not supported.")
	
	softwareIosXe = [ x.split('IOS XE ')[1] for x in softwareList if 'IOS XE' in x.upper() ]
	softwareSet = set()
	for item in softwareIosXe:
		softwareSet.add(item)
	softwareIosXe = list(softwareSet)
	# print(softwareIosXe)
	
	softwareIos = [ x.split('IOS ')[1] for x in softwareList if 'IOS' in x and 'IOS XE' not in x ]
	softwareSet = set()
	for item in softwareIos:
		softwareSet.add(item)
	softwareIos = list(softwareSet)	
	# print(softwareIos)
	
	softwareCantDo = [ x for x in softwareList if 'IOS' not in x ]
	softwareSet = set()
	for item in softwareCantDo:
		softwareSet.add(item)
	softwareCantDo = list(softwareSet)		
	# print(softwareCantDo)
	
	cantAnalyze = []
	for item in softwareCantDo:
		cantAnalyze += [[item]]
	# print(cantAnalyze)
	
	# extract the authentication information
	try:
		with open_file('auth.txt','r') as fileIn:
			[ ciscoAPIName, ciscoAPIPwd ] = [ x.rstrip() for x in fileIn.readlines() ]
	except:
		print("Required auth.txt file with info does not exist. Exiting...")
		sys.exit()
	
	# get auth token
	global accessToken	# this needs to stay global - may need to be refreshed during the script run
	global headers		# this needs to stay global - may need to be refreshed during the script run
	accessToken, headers = getAuthToken(ciscoAPIName, ciscoAPIPwd)

	# define the URLs for IOS and IOSXE
	iosURL = "https://api.cisco.com/security/advisories/ios?version="
	iosxeURL = "https://api.cisco.com/security/advisories/iosxe?version="
	
	# initialize some dictionaries for cases of no vulnerabilities detected or IOS not found
	noVulnFound = []
	iosNotFound = []
	
	# get the IOS vulnerabilities
	colHeaders = []
	outExcel = []	
	for iosversion in softwareIos:
		print("gathering information on: " + iosversion)
		response = requests.get(iosURL + iosversion, headers = headers)
		try:
			outputDict = json.loads(response.text)
		except:
			# case where it is a string - maybe not authorized?
			if 'Not Authorized' in response.text:
				# if not authorized - then get a new token and repeat the request
				accessToken, headers = getAuthToken(ciscoAPIName, ciscoAPIPwd)
				response = requests.get(iosURL + iosversion, headers = headers)
				try:
					outputDict = json.loads(response.text)
				except:
					# this is not working
					print("\t\t\tThis IOS version is not working. Troubleshoot.")
					continue	

		# check to see if it responds with an error - might be an IOS XE - and can be added to that list
		if 'errorCode' in outputDict.keys():
			if outputDict['errorCode'] == 'INVALID_IOS_VERSION':
				print('\t\tNot found via IOS query. Will try IOS XE after IOS finishes processing.')
				# print(response.text)
				softwareIosXe += [iosversion]
				continue
			elif outputDict['errorCode'] == 'NO_DATA_FOUND':
				noVulnFound += ["IOS " + iosversion]
				continue
				
		# make a list from the dictionary and add it to an Excel file
		outputDict = json.loads(response.text)['advisories']
		
		if colHeaders == []:
			# if the first time - get the column headers and filter them - then set in the list to output to Excel
			colHeaders = list(outputDict[0].keys())
			colHeaders = [ x for x in colHeaders if x not in ['summary','productNames','ipsSignatures','cvrfUrl','ovalUrl','bugIDs','cves','cwe'] ]
			outExcel.append(colHeaders)

		# then get information on each vulnerability
		for openVuln in outputDict:
			newList = [ openVuln[x] for x in colHeaders ]
			newList2 = []
			for item in newList:
				if type(item) is list:
					newList2.append(','.join(item))
				else:
					newList2.append(item)
			outExcel.append(newList2)
	outputExcel(outExcel,outputExcelFile,"IOS Vulnerabilities")

	# get the IOS XE vulnerabilities
	colHeaders = []
	outExcel = []	
	for iosversion in softwareIosXe:
		print("gathering information on: " + iosversion)
		response = requests.get(iosxeURL + iosversion, headers = headers)
		try:
			outputDict = json.loads(response.text)
		except:
			# case where it is a string - maybe not authorized?
			if 'Not Authorized' in response.text:
				# if not authorized - then get a new token and repeat the request
				accessToken, headers = getAuthToken(ciscoAPIName, ciscoAPIPwd)
				response = requests.get(iosURL + iosversion, headers = headers)
				try:
					outputDict = json.loads(response.text)
				except:
					# this is not working
					print("\t\t\tThis IOS version is not working. Troubleshoot.")
					continue

		# check to see if it responds with an error - at that point no luck
		if 'errorCode' in outputDict.keys():
			if outputDict['errorCode'] == "INVALID_IOSXE_VERSION":
				print('\t\tNot found via any queries.')
				iosNotFound += [iosversion]
				continue
			elif outputDict['errorCode'] == 'NO_DATA_FOUND':
				noVulnFound += ["IOS XE " + iosversion]			
				continue

		# make a list from the dictionary and add it to an Excel file
		outputDict = json.loads(response.text)['advisories']
	
		if colHeaders == []:
			# if the first time - get the column headers and filter them - then set in the list to output to Excel
			colHeaders = list(outputDict[0].keys())
			colHeaders = [ x for x in colHeaders if x not in ['summary','productNames','ipsSignatures','cvrfUrl','ovalUrl','bugIDs','cves','cwe'] ]
			outExcel.append(colHeaders)

		# then get information on each vulnerability
		for openVuln in outputDict:
			newList = [ openVuln[x] for x in colHeaders ]
			newList2 = []
			for item in newList:
				if type(item) is list:
					newList2.append(','.join(item))
				else:
					newList2.append(item)
			outExcel.append(newList2)
	outputExcel(outExcel,outputExcelFile,"IOS XE Vulnerabilities")
	
	# output the 'error' information to additional tabs (if there is anything to export)
	try:
		outputExcel(noVulnFound,outputExcelFile,"No Vulnerabilities Found")
	except:
		pass
	try:
		outputExcel(iosNotFound,outputExcelFile,"Software Not Found")
	except:
		pass
	try:
		outputExcel(cantAnalyze,outputExcelFile,"No Tool Support")
	except:
		pass
	
if __name__ == "__main__":

	# this gets run if the script is called by itself from the command line
	main(sys.argv)