@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  Excel Local Lazy - Tao shortcut Desktop
echo ============================================
echo.

rem --- Find pythonw.exe (same order as start.cmd) ---
set "PYW="
if exist "C:\Python314\pythonw.exe" set "PYW=C:\Python314\pythonw.exe"
if not defined PYW (
  for /f "delims=" %%i in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%i"
)
if not defined PYW if exist "%LocalAppData%\Programs\Python\Python314\pythonw.exe" set "PYW=%LocalAppData%\Programs\Python\Python314\pythonw.exe"
if not defined PYW (
  echo [LOI] Khong tim thay pythonw.exe tren may.
  echo       Hay cai Python 3.14 hoac chay start.cmd de xem chi tiet.
  pause
  exit /b 1
)
echo Tim thay Python: %PYW%
echo Dang tao shortcut tren Desktop va Start Menu...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$targets = @([Environment]::GetFolderPath('Desktop'), (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'));" ^
  "foreach ($dir in $targets) {" ^
  "  $lnk = $ws.CreateShortcut((Join-Path $dir 'Excel Local Lazy.lnk'));" ^
  "  $lnk.TargetPath = '%PYW%';" ^
  "  $lnk.Arguments = '\"%~dp0app.py\"';" ^
  "  $lnk.WorkingDirectory = '%~dp0';" ^
  "  if (Test-Path '%~dp0icon.ico') { $lnk.IconLocation = '%~dp0icon.ico' };" ^
  "  $lnk.Description = 'Excel Local Lazy - Tro ly bang tinh cuc bo va AI';" ^
  "  $lnk.Save();" ^
  "  Write-Host ('Da tao: ' + (Join-Path $dir 'Excel Local Lazy.lnk'));" ^
  "}"

if errorlevel 1 (
  echo [LOI] Tao shortcut that bai. Chup man hinh nay gui lai de duoc ho tro.
  pause
  exit /b 1
)
echo.
echo Xong! Tu nay chi can bam dup bieu tuong "Excel Local Lazy"
echo tren Desktop (hoac trong Start Menu) la app mo ngay,
echo khong can vao thu muc, khong can chay file nao khac.
pause
