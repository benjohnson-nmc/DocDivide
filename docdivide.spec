# docdivide.spec
# PyInstaller spec file for DocDivide
#
# BEFORE BUILDING:
#   1. pip install pyinstaller anthropic pypdf2 pdf2image pillow ttkbootstrap
#   2. Download Poppler for Windows:
#      https://github.com/oschwartz10612/poppler-windows/releases
#      Extract to a folder, e.g. C:\poppler
#      Set POPPLER_PATH below to that folder's bin\ directory.
#   3. Run: pyinstaller docdivide.spec

import os
import sys
import tempfile
from pathlib import Path

# -- CONFIGURE THIS --
POPPLER_BIN = r"C:\Users\bjohnson\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin"
APP_ICON    = "icon.ico"                      # set to None if no icon
# ---------------------

# Write the runtime hook that puts bundled poppler on PATH
hook_code = """
import os, sys
poppler_path = os.path.join(sys._MEIPASS, "poppler")
os.environ["PATH"] = poppler_path + os.pathsep + os.environ.get("PATH", "")
"""
hook_path = Path(tempfile.gettempdir()) / "hook_poppler.py"
hook_path.write_text(hook_code)

poppler_bins = [(str(f), "poppler") for f in Path(POPPLER_BIN).glob("*.dll")]
poppler_bins += [(str(f), "poppler") for f in Path(POPPLER_BIN).glob("*.exe")]

a = Analysis(
    ["docdivide.py"],
    pathex=[],
    binaries=poppler_bins,
    datas=[],
    hiddenimports=[
        "anthropic",
        "pypdf",
        "PyPDF2",
        "pdf2image",
        "PIL",
        "PIL._tkinter_finder",
        "ttkbootstrap",
        "tkinter",
        "tkinter.ttk",
        "tkinter.filedialog",
        "tkinter.messagebox",
        "oracledb",
        "tkinterdnd2",
    ],
    hookspath=[],
    runtime_hooks=[str(hook_path)],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DocDivide",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=APP_ICON if APP_ICON and Path(APP_ICON).exists() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="DocDivide",
)
