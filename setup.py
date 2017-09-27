import sys
from cx_Freeze import setup, Executable

setup(
    name = "AutomateMcMaster",
    version = "1.2",
    description = "An easier way to tabulate POs",
    executables = [Executable("AutomateMcMaster_2016-10-03.py", base = "Win32GUI")])