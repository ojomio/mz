import sys
from distutils.core import setup

from cx_Freeze import Executable, setup

# Dependencies are automatically detected, but it might need fine tuning.
# build_exe_options = {"packages": ["os"], "excludes": ["tkinter"]}
build_exe_options = {}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Printer",
    version="0.1",
    description="Printtr",
    options={"build_exe": build_exe_options},
    executables=[Executable("print.py", base=base, icon='Icon.ico')],
)
