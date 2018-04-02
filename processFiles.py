import os, sys, time, re, datetime
import pyautogui
import win32gui
import win32com.client
from win32com.client import Dispatch

def processFiles():
	#create MonthlyLoadLog Folder if needed
	if getattr(sys, 'frozen', False):
	    # frozen
	    dir_path_tmp = os.path.dirname(sys.executable).split("\\")
	else:
	    # unfrozen
	    dir_path_tmp = os.path.dirname(os.path.realpath(__file__)).split("\\")

	dir_path_root = "\\".join(dir_path_tmp[0:3])
	dir_path=dir_path_root+"\\Documents\\Analytics\\MonthlyLoads"
	monthlyLoadLogDir = dir_path+"\logs"

	#make sure Analytics / monthly loads directory is available.
	if not os.path.exists(dir_path):
		os.makedirs(dir_path)
		print("Created a MonthlyLoads directory.")

	if not os.path.exists(monthlyLoadLogDir):
	    os.makedirs(monthlyLoadLogDir)
	    print("created logging folder.")

	#import dates/times for filenames
	now = datetime.datetime.now()
	year = now.year
	month = '%02d' % now.month
	day = now.day
	hour=now.hour
	minute=now.minute
	second=now.second
	fullDate = [year, month, day]
	fullDateTime = [year,month,day,hour,minute,second]
	processingDate = "_".join(map(str, fullDate))
	processingDateTime = "_".join(map(str, fullDateTime))

	#make log for the monthly load itself
	logFile= open(monthlyLoadLogDir+"\\"+"Processing_"+processingDateTime+".txt","w+")
	logFile.write("File processing started.\n")
	#define xlApp globally
	xlApp = win32com.client.Dispatch("Excel.Application")

	#define function to open excel and jobpatch-scripts.xla
	def openExcel():
		xlApp = win32com.client.Dispatch("Excel.Application")
		xlApp.Visible = True

		try:
			xlApp.Workbooks.Open(dir_path+"\jobpatch-scripts.xla")
			logFile.write("Excel opened, job patch scripts loaded.\n")
			print("Excel opened, job patch scripts loaded.")
		except Exception as error:	
			print(str(error))
			logFile.write("Problem loading jobpatch-scripts.xla. Please see error:"+ str(error)+"\n")

	#Open the xla addin / vba scripts to make them available at runtime
	openExcel()

	#set monthly loads path to open / work with the downloaded data files.
	path = dir_path+"\dataFiles"

	#set ATS and JobPatch Regex up
	jobPatchRegex = re.compile("\d*_[a-zA-Z&]*_(JOB|JOBPatchExtract)_\d*.csv")
	atsRegex = re.compile("\d*_[a-zA-Z&]*_A.*_.*.(csv|CSV)")

	#start count of files to process
	numProcessed = 0
	#total number of files to process
	filesToProcess=len(os.listdir(path))
	#parse filename for scripting
	for file in os.listdir(path):
		#close and reopen excel every 10 files to limit memory usage errors.
		if numProcessed % 10 == 0 and numProcessed != 0:
			xlApp.Quit()
			logFile.write("Processed "+ str(numProcessed)+" files. Excel closed for performance.\n")
			print("Processed "+ str(numProcessed)+" files. Excel closed for performance.")
			openExcel()

		logFile.write("\n"+"\n")
		filename = os.fsdecode(file)	

		tempID = filename.split("_")
		siteID = tempID[0]
		
		#clean name if it has & in it
		if "&" in tempID[1]:
			name= tempID[1].replace("&", "and")
		
		else: name = tempID[1]
		
		siteIDName = siteID+"_"+name
		
		#Open the file we want in Excel
		try:
			workbook = xlApp.Workbooks.Open(path+"\\"+filename)
			logFile.write(filename+" successfully opened. \n")
			print(filename+" successfully opened.")
		except Exception as error:
			print(str(error))
			logFile.write("Problem loading"+filename+". Please see error: "+str(error)+"\n")	

	#Process job patches
		if jobPatchRegex.match(filename):
			#Open the file we want in Excel
			try:
				xlApp.Run(name+"_JOB")
				logFile.write("Jobpatch script run successfully against "+filename+".\n")
				print("Jobpatch script run successfully against "+filename)
				
				excelOutPath=path+"\\"+siteIDName+"_JOB_"+processingDate+".txt"
				
				try: 
					workbook.SaveAs(excelOutPath, FileFormat=42)
					workbook.Close()
					logFile.write(filename+" saved successfully.\n")
					print(filename+" saved successfully.")
				
				except Exception as error:
					print(str(error))
					logFile.write("Unable to save "+filename+" successfully. "+str(error)+"\n")	
			
			except Exception as error:
				print(str(error))
				logFile.write("Error running jobpatchscript on "+filename+". Please see error: "+str(error)+".\n")

			#process the resulting .txt file in Notepad to remove ".	
			shell = win32com.client.Dispatch("WScript.Shell")
			try:
				shell.Run("notepad.exe " + excelOutPath)
				logFile.write("Opened "+filename+".txt file in notepad. \n")
				print("Opened "+filename+".txt file in notepad.")
				
				#slight 1 second pause so notepad can catch up and become the active window.
				time.sleep(1)
				#Check to make sure active window is notepad, wait until notepad is active window / focus to continue.
				notepadTitleRegex = re.compile(".* - Notepad")
				window = win32gui.GetWindowText(win32gui.GetForegroundWindow())
				while notepadTitleRegex.match(window) is None:
					time.sleep(1)
				
				else:
					#open Replace
					pyautogui.hotkey("ctrl", "h")
					#type in the quote symbol
					pyautogui.press('"')
					#Tab down 5 times to the Replace All button
					pyautogui.press("tab")
					pyautogui.press("tab")
					pyautogui.press("tab")
					pyautogui.press("tab")
					pyautogui.press("tab")
					#Press enter to remove all quotes
					pyautogui.press("enter")
					#tab down to cancel to get out of the replace all window.
					pyautogui.press("tab")
					pyautogui.press("enter")
					#save the file
					pyautogui.hotkey("ctrl", "s")
					time.sleep(.5)
					shell.run("tskill notepad")
					
					logFile.write("File processed / quotes removed successfully. \n")
					print("File processed / quotes removed successfully.")
					
					numProcessed+=1

			except Exception as error:
				print(str(error))
				logFile.write("Unable to open .txt file for "+filename+". "+str(error)+"\n")	
			
	#Process ATS scripts 
		if atsRegex.match(filename):
			#open the file we want in Excel
			try:
				xlApp.Run(name+"_ATS")
				logFile.write("ATS script run successfully against "+filename+".\n")
				print("ATS script run successfully against "+filename+".")
				
				excelOutPath=path+"\\"+siteIDName+"_ATS_"+processingDate+".xlsx"
				
				try:
					workbook.SaveAs(excelOutPath, FileFormat=51)
					workbook.Close()
					logFile.write(filename+" saved successfully.\n")
					print(filename+" saved successfully.")
					
					numProcessed+=1
				
				except Exception as error:
					print(str(error))
					logFile.write("Unable to save "+filename+" successfully. "+str(error)+"\n")	
			
			except Exception as error:
				print(str(error))
				logFile.write("Error running jobpatchscript on "+filename+". Please see error: "+str(error)+".\n")

	xlApp.Quit()
	print("\n")
	print("Monthly load processing complete.")

if __name__ == "__main__":
	processFiles()
