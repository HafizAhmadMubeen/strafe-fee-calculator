@echo off
REM ============================================================================
REM  Build  dist\STRAFE Fee Calculator.exe  - one file, no Python needed to run.
REM  Run this once on a machine that has the .venv set up.
REM ============================================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Setup not done. First run:
  echo    python -m venv .venv
  echo    .venv\Scripts\python -m pip install playwright openpyxl
  pause
  exit /b 1
)

echo Installing PyInstaller into the venv...
".venv\Scripts\python.exe" -m pip install --upgrade pyinstaller || (
  echo pip install pyinstaller failed - ask IT to allow it, or install it manually.
  pause
  exit /b 1
)

echo.
echo Building (this takes a few minutes and the .exe will be ~150-200 MB)...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm "STRAFE Fee Calculator.spec" || (
  echo Build failed - see the messages above.
  pause
  exit /b 1
)

echo.
echo ============================================================================
echo  Done:  "%~dp0dist\STRAFE Fee Calculator.exe"
echo  Give that ONE file to users. First launch is slow (it unpacks itself).
echo  The machine running it needs Microsoft Edge installed.
echo ============================================================================
pause