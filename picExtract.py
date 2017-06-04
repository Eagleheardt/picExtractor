#!/usr/bin/python3

import shutil
import os
import zipfile

def zipDoc(aFile,dirPath):
	dotNDX = aFile.index(".") # position of the .
	shortFN = aFile[:dotNDX] # name of the file before .
	zipName = dirPath + shortFN + ".zip" # name and path of the file only .zip
	shutil.copy2(dirPath + aFile, zipName) # copies all data from original into .zip format
	useZIP = zipfile.ZipFile(zipName) # the usable zip file
	return useZIP

def hasPicExtension(aFile): # if a file ends in a typical picture file extension, returns true
	if(aFile.endswith(".jpeg") or aFile.endswith(".jpg") or aFile.endswith(".png") or aFile.endswith(".JPEG") or aFile.endswith(".JPG") or aFile.endswith(".PNG") or aFile.endswith(".bmp") or aFile.endswith(".BMP")):
		return True
	else:
		return False

def delDOCXEvidence(somePath):
	##################################################################
	# Working Linux code:
	os.rmdir(somePath + "/word/media") # removes directory
	os.rmdir(somePath + "/word") # removes more directory
	##################################################################
				
	##################################################################
	# Untested windows code:
	# os.rmdir(somePath + "\\\\word\\\\media") # removes directory
	# os.rmdir(somePath + "\\\\word") #removes more directory
	##################################################################

def delXLSXEvidence(somePath):
	##################################################################
	# Working Linux code:
	os.rmdir(somePath + "/xl/media") # removes directory
	os.rmdir(somePath + "/xl") # removes more directory
	##################################################################
				
	##################################################################
	# Untested windows code:
	# os.rmdir(somePath + "\\\\word\\\\media") # removes directory
	# os.rmdir(somePath + "\\\\word") #removes more directory
	##################################################################

def extractPicsFromDir(dirPath=""):
# when given a directory path, will extract all images from all .docx and .xlsx file types
	if os.path.isdir(dirPath): # if the given path is a directory
		for dirFile in os.listdir(dirPath): # loops through all files in the directory
			dirFileName = os.fsdecode(dirFile) # strips out the file name
			if dirFileName.endswith(".docx"):
				useZIP = zipDoc(dirFile,dirPath) # turns it into a zip
				picNum = 1 # number of pictures in file
				for zippedFile in useZIP.namelist(): # loops through all files in the directory
					if hasPicExtension(zippedFile): # if it ends with photo
						useZIP.extract(zippedFile, path=dirPath) # extracts the picture to the path + word/media/
						shutil.move(dirPath + str(zippedFile),dirPath + dirFileName[:dirFileName.index(".")] + " - " + str(picNum)) # moves the picture out
						picNum += 1
				delDOCXEvidence(dirPath) # removes the extracted file structure
				os.remove(useZIP.filename) # removes zip file
				# no evidence
			if dirFileName.endswith(".xlsx"):
				useZIP = zipDoc(dirFile,dirPath) # turns it into a zip
				picNum = 1 # number of pictures in file
				for zippedFile in useZIP.namelist(): # loops through all files in the directory
					if hasPicExtension(zippedFile): # if it ends with photo
						useZIP.extract(zippedFile, path=dirPath) # extracts the picture to the path + word/media/
						shutil.move(dirPath + str(zippedFile),dirPath + dirFileName[:dirFileName.index(".")] + " - " + str(picNum)) # moves the picture out
						picNum += 1
				delXLSXEvidence(dirPath) # removes the extracted file structure
				os.remove(useZIP.filename) # removes zip file
				# no evidence
				
	else:
		print("Not a directory path!")
		exit(1)


uDir = input("Enter your directory: ")
extractPicsFromDir(uDir)
