import os
import sys
import shutil
from dotenv import load_dotenv
from tkinter import *
from tkinter import messagebox

def processInput(userInput):
	print(userInput)
	networkStatus = directAccessReminder()
	if networkStatus == True:
		if userInput == "Monthly Download":
			syncScriptToLocal("monthly", "download")

		elif userInput == "Monthly Delete":
			syncScriptToLocal("monthly", "delete")

		elif userInput == "Weekly Delete":
			syncScriptToLocal("weekly", "delete")

		else:
			syncScriptToLocal("weekly", "download")

		print("File sync is complete.")
	else:
		print("Waiting for user to accept DirectAccess Reminder.")
		return

def directAccessReminder():
	#Tkinter setup
	networkStatus = messagebox.askquestion("DirectAccess Reminder", """You must be on DirectAccess to connect to SFTP server via WinSCP. 
	                   					Please disconnect from VPN prior to running WINSCP scripts.""",icon='warning')
	if networkStatus == "yes":
		return True
	else:
		return False		

def syncScriptToLocal(frequency, actionType):
	#set dir_path if frozen or normal python script
	if getattr(sys, 'frozen', False):
	    # frozen
	    dir_path_tmp = os.path.dirname(sys.executable).split("\\")
	else:
	    # unfrozen
	    dir_path_tmp = os.path.dirname(os.path.realpath(__file__)).split("\\")

	dir_path_root = "\\".join(dir_path_tmp[0:3])
	dir_path = dir_path_root+"\\Documents\\Analytics\\MonthlyLoads"

	#Dotenv variables
	dotenv_path = os.path.join(dir_path,".env")
	load_dotenv(dotenv_path)
	SHARED_DRIVE_SCRIPTING = os.path.join(os.environ.get("SHARED_DRIVE_SCRIPTING"),"MonthlyLoads/winscpScripts")

	#set shared drive paths
	monthlyDownloadShared = os.path.join(SHARED_DRIVE_SCRIPTING, "monthlyDownload.txt")
	monthlyDeleteShared = os.path.join(SHARED_DRIVE_SCRIPTING, "monthlyDelete.txt")
	weeklyDownloadShared = os.path.join(SHARED_DRIVE_SCRIPTING, "weeklyDownload.txt")
	weeklyDeleteShared = os.path.join(SHARED_DRIVE_SCRIPTING, "weeklyDelete.txt")

	#set local Analytics Path

	winSCPDirectory = os.path.join(dir_path,"winscpScripts")
	monthlyDownloadWinSCP = os.path.join(winSCPDirectory, "monthlyDownload.txt")
	monthlyDeleteWinSCP = os.path.join(winSCPDirectory, "monthlyDelete.txt")
	weeklyDownloadWinSCP = os.path.join(winSCPDirectory, "weeklyDownload.txt")
	weeklyDeleteWinSCP = os.path.join(winSCPDirectory, "weeklyDelete.txt")
	batchPath = os.path.join(winSCPDirectory, "batch.bat")

	script=""
	#set file paths
	if frequency == "monthly" and actionType == "download":
		shutil.copy(monthlyDownloadShared, monthlyDownloadWinSCP)
		command = batchPath + " monthlyDownload.txt"

	elif frequency == "monthly" and actionType == "delete":
		shutil.copy(monthlyDeleteShared, monthlyDeleteWinSCP)
		command = batchPath + " monthlyDelete.txt"

	elif frequency == "weekly" and actionType == "delete":
		shutil.copy(weeklyDeleteShared, weeklyDeleteWinSCP)
		command = batchPath + " weeklyDelete.txt"

	else: #frequency == "weekly" and actionType == "download"
		shutil.copy(weeklyDownloadShared, weeklyDownloadWinSCP)
		command = batchPath + " weeklyDownload.txt"

	#execute the batch file
	os.chdir(winSCPDirectory)
	os.system(command)	