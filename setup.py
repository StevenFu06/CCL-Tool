import sys, os
from cx_Freeze import setup, Executable
import matplotlib

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))

sys.path.append('pandastable')

# currently requires changing line 548 of hooks.py to make scipy work
# see https://bitbucket.org/anthony_tuininga/cx_freeze/issues/43

includefiles = ["pandastable/dataexplore.gif", "pandastable/datasets",
                "pandastable/plugins",
                os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
                os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
                'Poppler/',
                'Tesseract-OCR/',
                (matplotlib.get_data_path(), "mpl-data"),
                ]
packages = ['docx', 'selenium', 'pickle', 'os', 'time', 'pathlib', 'bs4', 'tkinter', 'pandastable', 'threading',
            'pandas', 're', 'zipfile', 'json', 'copy', 'shutil', 'io', 'concurrent', 'pdfminer', 'pytesseract',
            'pdf2image', 'matplotlib', 'numpy', 'mpl_toolkits', 'multiprocessing', 'StyleFrame']

options = {
    'build_exe': {
        'packages': packages,
    },
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("gui.py", base=base)]

setup(name="CCL Tool",
      options=options,
      version="1.0",
      description="Critical Components list Tool",
      executables=executables)
