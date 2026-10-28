@echo off
set PYTHONUTF8=1
cd /d "%~dp0"
if exist "C:\Python314\pythonw.exe" (
  start "" "C:\Python314\pythonw.exe" "%~dp0app.py"
) else (
  python "%~dp0app.py"
)
