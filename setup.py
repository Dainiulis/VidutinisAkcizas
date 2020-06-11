import os
from cx_Freeze import setup, Executable
import sys

os.environ['TCL_LIBRARY'] = r"C:\Users\daini\AppData\Local\Programs\Python\Python36\tcl\tcl8.6"
os.environ['TK_LIBRARY'] = r"C:\Users\daini\AppData\Local\Programs\Python\Python36\tcl\tk8.6"

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("gui_tkinter\\main_window.py", base=base)]

packages = ["os", "tkinter", 'numpy']
include_files = [
    r"C:\Users\daini\AppData\Local\Programs\Python\Python36\DLLs\tcl86t.dll",
    r"C:\Users\daini\AppData\Local\Programs\Python\Python36\DLLs\tk86t.dll"
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
    version="1.05",
    description='Mini pakeitimai',
    executables=executables
)
