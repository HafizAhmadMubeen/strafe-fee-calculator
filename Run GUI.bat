@echo off
REM Double-click to open the STRAFE Fee Calculator window.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Setup not done yet. Open a terminal here and run:
  echo    python -m venv .venv
  echo    .venv\Scripts\python -m pip install playwright openpyxl
  pause
  exit /b 1
)
".venv\Scripts\pythonw.exe" fill_calculator_gui.py