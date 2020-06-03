import sys, os
from cx_Freeze import setup, Executable

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))

sys.path.append('pandastable')

# currently requires changing line 548 of hooks.py to make scipy work
# see https://bitbucket.org/anthony_tuininga/cx_freeze/issues/43

includes = ["pandastable"]
includefiles = ["pandastable/dataexplore.gif", "pandastable/datasets",
                "pandastable/plugins",
                os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
                os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll')
                ]

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("gui.py", base=base,
                          targetName='DataExplore.exe',
                          shortcutName="DataExplore",
                          shortcutDir="DesktopFolder",
                          icon="img/dataexplore.ico")]

setup(name="DataExplore",
      version="0.12.2",
      description="Data analysis and plotter",
      executables=executables)
