from tkinter import *
from tkinter import messagebox
import sys
import os
import dotenv
import re
import shutil
import pyautogui
import win32gui
import win32com.client
from win32com.client import Dispatch


#path to module scripts
#detemine if frozen or not
if getattr(sys, 'frozen', False):
    # frozen
    dir_path_tmp = os.path.dirname(sys.executable).split("\\")
else:
    # unfrozen
    dir_path_tmp = os.path.dirname(os.path.realpath(__file__)).split("\\")
dir_path_root = "\\".join(dir_path_tmp[0:3])
dir_path=dir_path_root+"\\Documents\\Analytics\\MonthlyLoads"
#make other modules available even when running as exe from build directory.
sys.path.insert(0, dir_path)

import sftpfunctions as sftpfunctions
import processFiles as processFiles
import moveFiles as moveFiles

#Tkinter setup
inputBox = Tk()

inputBox.title("Analytics File Processing")
leftFrame = Frame(inputBox)
leftFrame.pack(side=LEFT, fill=X, anchor=W,padx=20,pady=20)

titleLabel = Label(leftFrame, font=('arial', 11, 'bold'),
                   text="SFTP File Download or Deletion",
                   bd=5, anchor=W)
titleLabel.pack(anchor=NW)
var = StringVar(inputBox)
var.set("Select Option") # initial value


selectLabel = Label(leftFrame, text="Please select a processing option below.")
selectLabel.pack(anchor=NW, pady = 15)
option = OptionMenu(leftFrame, var, "Monthly Download", "Monthly Delete", "Weekly Download", "Weekly Delete")
option.pack(side=LEFT, ipadx=30,ipady=5)
#arrow = PhotoImage(file='down.png')
option.configure(indicatoron=0, compound='right')
okOption = Button(leftFrame, text="Run", command= lambda: sftpfunctions.processInput(var.get()))
okOption.pack(side=LEFT,ipadx=30,ipady=6)

processFilesBtn = Button(leftFrame, text="Process Files", command=processFiles.processFiles)
processFilesBtn.pack(side=LEFT, padx=(35,20), ipadx=30,ipady=6)

moveFilesBtn = Button(leftFrame, text="Move Files", command=moveFiles.moveFiles)
moveFilesBtn.pack(side=LEFT, ipadx=30,ipady=6)
userInput = var.get()

inputBox.mainloop()
