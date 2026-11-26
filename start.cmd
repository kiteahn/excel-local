@echo off
set PYTHONUTF8=1
cd /d "%~dp0"
rem Excel Local Lazy launcher: prefer pythonw (no console window).
rem Detection order keeps the known-good machine path first.
if exist "C:\Python314\pythonw.exe" (
  start "" "C:\Python314\pythonw.exe" "%~dp0app.py"
  exit /b 0
)
for /f "delims=" %%i in ('where pythonw 2^>nul') do (
  start "" "%%i" "%~dp0app.py"
  exit /b 0
)
if exist "%LocalAppData%\Programs\Python\Python314\pythonw.exe" (
  start "" "%LocalAppData%\Programs\Python\Python314\pythonw.exe" "%~dp0app.py"
  exit /b 0
)
rem Fallback: console python (window stays visible, but the app still runs).
python "%~dp0app.py"
