@echo off
REM ============================================================================
REM  Launch Microsoft Edge with the remote-debugging port that fill_calculator.py
REM  needs to attach to.  Double-click this file (or run it from a terminal).
REM
REM  IMPORTANT: close ALL other Edge windows first, otherwise Windows just opens
REM  a new tab in the running Edge and the debugging port is NOT enabled.
REM
REM  This uses a dedicated Edge profile folder (.edge-automation, next to this
REM  file) so it never disturbs your normal Edge profile.  The first time, you
REM  log in to STRAFE in the window it opens; the login is remembered in that
REM  folder for later runs.
REM ============================================================================

set "PROFILE=%~dp0.edge-automation"
set "URL=https://strafe.stewart.com/TestingCalculator"

set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" (
  echo Could not find msedge.exe in the usual locations.
  echo Edit this .bat and set EDGE to the full path of msedge.exe.
  pause
  exit /b 1
)

echo Launching Edge with remote debugging on port 9222...
start "" "%EDGE%" --remote-debugging-port=9222 --user-data-dir="%PROFILE%" "%URL%"

echo.
echo Edge is starting. In that window:
echo   1. Log in to STRAFE if prompted.
echo   2. Make sure the Testing Calculator page is open.
echo   3. Leave it on that page - do not navigate away.
echo.
echo Then, in the VS Code terminal, run:
echo   .venv\Scripts\python fill_calculator.py "202x xx xx UT STGC Ends - Test Sheet -1.xlsx" --limit 3
echo.