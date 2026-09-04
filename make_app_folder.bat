@echo off
REM ============================================================================
REM  Assembles a single flat folder ready to push to the drive:
REM     "STRAFE Fee Calculator\"  containing just the exe + the user guide
REM     (+ launch_edge_debug.bat, needed for the "Connect Edge" button to use
REM      the IT-approved launcher).
REM  For the fuller package (with Source/docs/Sample workbook) use
REM  make_drive_package.bat instead.
REM ============================================================================
setlocal
cd /d "%~dp0"

set "OUT=STRAFE Fee Calculator"

if not exist "dist\STRAFE Fee Calculator.exe" (
  echo dist\STRAFE Fee Calculator.exe not found - run build_exe.bat first.
  pause
  exit /b 1
)

echo Rebuilding the User Guide PDF...
".venv\Scripts\python.exe" "docs\md_to_pdf.py" "docs\User Guide.md" >nul
if not exist "docs\User Guide.pdf" (
  echo Could not build docs\User Guide.pdf - see the error above.
  pause
  exit /b 1
)

echo Assembling "%OUT%" ...
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"
copy /y "dist\STRAFE Fee Calculator.exe" "%OUT%\" >nul
copy /y "docs\User Guide.pdf"            "%OUT%\" >nul
copy /y "launch_edge_debug.bat"          "%OUT%\" >nul

echo.
echo ============================================================================
echo  Done:  "%~dp0%OUT%"
echo  Upload that folder to the drive as-is.
echo ============================================================================
pause
