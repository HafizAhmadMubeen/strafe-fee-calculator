# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for the single-file "STRAFE Fee Calculator.exe".
#
# Build:
#     .venv\Scripts\python -m pip install pyinstaller
#     .venv\Scripts\python -m PyInstaller --noconfirm "STRAFE Fee Calculator.spec"
#
# Output: dist\STRAFE Fee Calculator.exe   (one file, no Python needed to run it)
#
# Notes
#  - It bundles the Playwright *Node driver* (needed even though we only attach to
#    an existing Edge over CDP).  It does NOT bundle any browser - the .exe drives
#    the user's installed Microsoft Edge.
#  - Windowed build (no console).  A startup crash is written to
#    %LOCALAPPDATA%\STRAFE Fee Calculator\last-error.log and shown in a dialog.

from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []
for pkg in ("playwright", "openpyxl", "et_xmlfile"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

hiddenimports += ["fill_calculator"]

a = Analysis(
    ["fill_calculator_gui.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "pandas", "numpy", "matplotlib"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="STRAFE Fee Calculator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # GUI app - no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="app.ico",
)