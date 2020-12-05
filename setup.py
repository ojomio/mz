import sys
from distutils.core import setup

from cx_Freeze import Executable, setup

# Dependencies are automatically detected, but it might need fine tuning.
# build_exe_options = {"packages": ["os"], "excludes": ["tkinter"]}
build_exe_options = {"excludes": ["tcl", "tkinter", "asyncio", "unittest", "email"]}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Printer",
    version="0.1",
    description="Подготовка распечаток для Wildberries",
    options={"build_exe": build_exe_options},
    packages=['qt'],
    executables=[Executable("print.py", base=base, icon='Icon.ico')],
)
