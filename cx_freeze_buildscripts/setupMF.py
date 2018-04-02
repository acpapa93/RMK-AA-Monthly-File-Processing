#CX_freeze script to build the MoveFiles.py into exe
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
					["tkinter","os","sys","time","datetime","re","shutil","dotenv"],
					"excludes": ["numpy"]}

setup(
	name="moveFiles",
	version="0.2",
	description="Move Jobpatch and ATS files to appropriate directory on shared drive.",
	executables=[Executable("moveFiles.py")])