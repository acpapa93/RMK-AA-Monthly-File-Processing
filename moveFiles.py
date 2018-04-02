import os
import sys
import time
import datetime
import re
from dotenv import load_dotenv
from shutil import copy

def moveFiles():
	#detemine if frozen or not
	if getattr(sys, 'frozen', False):
	    # frozen
	    dir_path_tmp = os.path.dirname(sys.executable).split("\\")
	else:
	    # unfrozen
	    dir_path_tmp = os.path.dirname(os.path.realpath(__file__)).split("\\")
	dir_path_root = "\\".join(dir_path_tmp[0:3])
	dir_path=dir_path_root+"\\Documents\\Analytics\\MonthlyLoads"

	#load up shared drive address / path
	dotenv_path = os.path.join(dir_path,".env")
	load_dotenv(dotenv_path)
	sharedDrive = os.environ.get("SHARED_DRIVE")
	correctionRegex = re.compile(os.environ.get("CORRECTION_REGEX"))
	correctedFileName = os.environ.get("CORRECTED_FILENAME")

	monthlyLoadLogDir = dir_path+"\logs"
	path = dir_path+"\dataFiles"

	#make sure Analytics / monthly loads directory is available.
	if not os.path.exists(dir_path):
		os.makedirs(dir_path)
		print("Created MonthlyLoads folder.")

	#create log Folder if needed
	if not os.path.exists(monthlyLoadLogDir):
	    os.makedirs(monthlyLoadLogDir)
	    print("Created monthly load log folder.")


	#import dates/times for filenames
	now = datetime.datetime.now()
	year = str(now.year)
	month = '%02d' % now.month
	hour=now.hour
	day=now.day
	minute=now.minute
	second=now.second
	monthName = now.strftime("%B")
	monthDirName = month+" - "+monthName
	ymd = [year,month,day]
	ymdJoined = "_".join(map(str, ymd))
	fullDateTime = [year,month,day,hour,minute,second]
	processingDateTime = "_".join(map(str, fullDateTime))

	#Regex to figure out if Analytics or JobPatch
	jobPatchRegex = re.compile("\d*_[a-zA-Z&]*_(JOB|JOBPatchExtract)_.*.(csv|txt|xlsx)")
	atsRegex = re.compile("\d*_[a-zA-Z&]*_A.*_.*.(csv|CSV|xlsx|xls)")

	#make log for the monthly load itself
	logFile= open(monthlyLoadLogDir+"\\"+"Moving_"+processingDateTime+".txt","w+")
	logFile.write("File movement started.\n")

	#total number of files to process
	filesToProcess=str(len(os.listdir(path)))

	for file in os.listdir(path):
		#set up logging breaks in between files
		processingTime = time.asctime()
		logFile.write("\n"+"\n")
		#set up file specific paths
		filename = os.fsdecode(file)
		srcPath = path+"\\"+filename

		#correct filename for one specific filename
		if correctionRegex.match(filename):
			suffix = filename.split(".")
			correctedName = correctedFileName+"_JOB_"+ymdJoined+"."+suffix[1]
			os.rename(path+"//"+filename, path+"//"+correctedName)
			filename = os.fsdecode(file)
			srcPath = path+"\\"+filename

		#get file name ready for client name + id
		tempClientID = filename.split("_")
		clientID = tempClientID[0]
		#clean clientName if it has & in it
		if "&" in tempClientID[1]:
			clientName= tempClientID[1].replace("&", "and")
		else: clientName = tempClientID[1]
		clientNameID = clientName+" - "+clientID
		clientPath = sharedDrive+"\\"+clientNameID
		#set logging init
		clientIDName = clientID+" - "+clientName

		#determine if jobpatch or Analytics file
		if jobPatchRegex.match(filename):
			yearPath = clientPath+"\\JobPatch\\"+year
			print(filename+" is a jobpatch")
			logFile.write(filename+" is a jobpatch. \n")

		if atsRegex.match(filename):
			yearPath = clientPath+"\\Analytics\\"+year
			print(filename+" is an ATS")
			logFile.write(filename+" is an ATS file.\n")

		#make year path if it doesn't exists
		if not os.path.exists(yearPath):
			os.makedirs(yearPath)
			print("Created YEAR directory within customer folder on Shared Drive.")
			logFile.write("Created YEAR directory within customer folder on Shared Drive.\n")

		#make sure month dir exists	
		destPath = yearPath+"\\"+monthDirName
		if not os.path.exists(destPath):
			os.makedirs(destPath)
			print("Created MONTH directory within customer folder on Shared Drive.")
			logFile.write("Created MONTH directory within customer folder on Shared Drive.\n")
		try:
			copy(srcPath, destPath)
			print(filename+" moved to "+destPath+"\n")
			logFile.write(filename+" moved to "+destPath+". \n")
		except Exception as error:
			print("Unable to move "+filename+" to "+destPath+". Error: "+error+"\n")
			logFile.write("Unable to move "+filename+" to "+ destPath+". Error: "+error+".\n")

	print("File move is complete. Please review logs for an errors")
	logFile.write("File move is complete. Please review logs for any errors")

if __name__ == "__main__":
	moveFiles()