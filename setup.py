import os
from cx_Freeze import setup, Executable
import sys

os.environ['TCL_LIBRARY'] = "C:/Users/daini/AppData/Local/Programs/Python/Python36/tcl/tcl8.6"
os.environ['TK_LIBRARY'] = "C:/Users/daini/AppData/Local/Programs/Python/Python36/tcl/tk8.6"

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("gui_tkinter\\main_window.py", base=base)]

packages = ["os", "tkinter", 'numpy']
include_files = [
    "C:/Users/daini/AppData/Local/Programs/Python/Python36/DLLs/tcl86t.dll",
    "C:/Users/daini/AppData/Local/Programs/Python/Python36/DLLs/tk86t.dll"
]
options = {
    'build_exe': {
        'packages': packages,
        "include_files": include_files
        },

}

setup(
    name="Vidutinis akcizas",
    options=options,
    version="1.00",
    description='Mini pakeitimai',
    executables=executables
)
