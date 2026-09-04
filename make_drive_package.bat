@echo off
REM ============================================================================
REM  Assembles  "drive package\STRAFE Fee Calculator\"  ready to upload to the
REM  shared drive.  Includes the built .exe if  dist\  has one.
REM ============================================================================
setlocal
cd /d "%~dp0"

set "OUT=drive package\STRAFE Fee Calculator"

echo Cleaning %OUT% ...
if exist "drive package" rmdir /s /q "drive package"
if exist "drive package" (
  echo   ! could not fully remove the old "drive package" folder - a file may
  echo     still be open ^(antivirus scanning it, Explorer preview, etc^). Close
  echo     anything using it and re-run this.
  pause
  exit /b 1
)
mkdir "%OUT%\App"
mkdir "%OUT%\Source"
mkdir "%OUT%\Sample workbook"
mkdir "%OUT%\docs"

REM ---- top-level readme --------------------------------------------------
copy /y "docs\drive-readme.txt" "%OUT%\README.txt" >nul

REM ---- App\  (for end users) ------------------------------------------
if exist "dist\STRAFE Fee Calculator.exe" (
  copy /y "dist\STRAFE Fee Calculator.exe" "%OUT%\App\" >nul
  if exist "%OUT%\App\EXE MISSING - build it first.txt" del "%OUT%\App\EXE MISSING - build it first.txt"
  echo   + exe added
) else (
  echo   ! dist\STRAFE Fee Calculator.exe not found - run build_exe.bat first.
  echo     Package will still be made, without the exe.
  echo Build the exe with build_exe.bat, then re-run this. > "%OUT%\App\EXE MISSING - build it first.txt"
)
copy /y "launch_edge_debug.bat" "%OUT%\App\" >nul
if exist "docs\User Guide.pdf" (
  copy /y "docs\User Guide.pdf" "%OUT%\App\" >nul
) else (
  copy /y "docs\User Guide.md" "%OUT%\App\User Guide.md" >nul
  echo   ! docs\User Guide.pdf not found - copied the .md instead.
  echo     Make the PDF with:  .venv\Scripts\python docs\md_to_pdf.py "docs\User Guide.md"
)

REM ---- Source\  (for the maintainer) --------------------------------
for %%F in (fill_calculator.py fill_calculator_gui.py build_exe.bat "STRAFE Fee Calculator.spec" "Run GUI.bat" requirements.txt README.md planning.md) do (
  copy /y "%%~F" "%OUT%\Source\" >nul
)

REM ---- Sample workbook\ ---------------------------------------------
for %%F in ("202x xx xx UT STGC Ends - Test Sheet -1.xlsx") do (
  if exist "%%~F" copy /y "%%~F" "%OUT%\Sample workbook\" >nul
)

REM ---- docs\ -----------------------------------------------------
copy /y "docs\User Guide.md"             "%OUT%\docs\" >nul
copy /y "docs\Build and Deploy Guide.md" "%OUT%\docs\" >nul
copy /y "docs\md_to_pdf.py"              "%OUT%\docs\" >nul
if exist "docs\User Guide.pdf" copy /y "docs\User Guide.pdf" "%OUT%\docs\" >nul

echo.
echo ============================================================================
echo  Done:  "%~dp0%OUT%"
echo  Upload the whole "STRAFE Fee Calculator" folder to the drive.
echo  Point users at the  App\  folder only.
echo ============================================================================
pause
