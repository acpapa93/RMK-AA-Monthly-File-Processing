#CX_freeze script to build the main.py into exe
import os
import sys
from cx_Freeze import setup, Executable

os.environ["TCL_LIBRARY"] = r"C:\Users\I856620\AppData\Local\Programs\Python\Python36\tcl\tcl8.6"
os.environ["TK_LIBRARY"] = r"C:\Users\I856620\AppData\Local\Programs\Python\Python36\tcl\tk8.6"

#detemine if frozen or not
if getattr(sys, 'frozen', False):
    # frozen
    dir_path_tmp = os.path.dirname(sys.executable).split("\\")
else:
    # unfrozen
    dir_path_tmp = os.path.dirname(os.path.realpath(__file__)).split("\\")
dir_path_root = "\\".join(dir_path_tmp[0:3])
dir_path=dir_path_root+"\\Documents\\Analytics\\MonthlyLoads"

#change into the root monthlyLoads directory.
os.chdir(dir_path)

build_exe_options = {"packages":
					["tkinter","os","sys","time","datetime","re","shutil","pyautogui","win32gui","win32com","dotenv"],
					"excludes": ["numpy"],
					"include_files": ["tcl86t.dll", "tk86t.dll"]}

setup(
	name="mainGUI",
	version="0.2",
	description="Process jobpatch files and ATS files.",
	executables=[Executable("mainGUI.py")],
	data_files=[("",["jobpatch-scripts.xla", ".env","README.txt",".gitignore"]), ("build_mainGui_amd64_3.6",["build_mainGui_amd64_3.6/tcl86t.dll","build_mainGui_amd64_3.6/tk86t.dll"]),("winscpScripts",["winscpScripts/batch.bat", "winscpScripts/winscp.com", "winscpScripts/winscpnet.dll", "winscpScripts/winscp.map", "winscpScripts/winscp.exe"])])