#!/usr/bin/python3

import shutil
import os
import zipfile

def extractPicsFromDir(dirPath=""):
# when given a directory path, will extract all images from all .docx and .xlsx file types
	if os.path.isdir(dirPath): # if the given path is a directory
		picNum = 1 # number of pictures
		for dirFile in os.listdir(dirPath): # loops through all files in the directory
			dirFileName = os.fsdecode(dirFile) # strips out the file name
			if dirFileName.endswith(".docx") or dirFileName.endswith(".xlsx"):
				dotNDX = dirFileName.index(".") # position of the .
				shortFN = dirFileName[:dotNDX] # name of the file before .
				zipName = dirPath + shortFN + ".zip" # name and path of the file only .zip
				shutil.copy2(dirPath + dirFileName, zipName) # copies all data from original into .zip format
				useZIP = zipfile.ZipFile(zipName) # the usable zip file
				for zippedFile in useZIP.namelist(): # loops through all files in the directory
					if zippedFile.endswith(".jpeg") or zippedFile.endswith(".jpg"): # if it ends with photo
						useZIP.extract(zippedFile, path=dirPath) # extracts the picture to the path + word/media/
						shutil.move(dirPath + str(zippedFile),"Picture - " + str(picNum)) # moves the picture out
						picNum += 1
				os.rmdir(dirPath + "/word/media") # removes directory
				os.rmdir(dirPath + "/word") # removes more directory
				##################################################################
				# Untested windows code:
				# os.rmdir(dirPath + "\\\\word\\\\media") # removes directory
				# os.rmdir(dirPath + "\\\\word") #removes more directory
				# os.remove(zipName) # removes zip file
				##################################################################
				os.remove(zipName) # removes zip file
				# no evidence
	else:
		print("Not a directory path!")
		exit(1)


uDir = input("Enter your directory: ")
extractPicsFromDir(uDir)
