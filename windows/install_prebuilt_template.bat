@echo off
setlocal
cd /d "%~dp0"

set "TARGET=%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm"
set "SOURCE=%~dp0Zotero.dotm"
set "BACKUP_DIR=%~dp0backup"
set "BACKUP=%BACKUP_DIR%\Zotero.backup.before-linking.dotm"

echo Please close Microsoft Word before install.
tasklist /FI "IMAGENAME eq WINWORD.EXE" | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo.
  echo Install failed: Microsoft Word is still running.
  pause
  exit /b 1
)

if not exist "%SOURCE%" (
  echo.
  echo Install failed: prebuilt Zotero.dotm not found.
  pause
  exit /b 1
)

if not exist "%APPDATA%\Microsoft\Word\STARTUP" (
  echo.
  echo Install failed: Word STARTUP folder was not found.
  echo Expected:
  echo %APPDATA%\Microsoft\Word\STARTUP
  pause
  exit /b 1
)

if exist "%TARGET%" (
  if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
  copy /y "%TARGET%" "%BACKUP%" >nul
)

copy /y "%SOURCE%" "%TARGET%" >nul
if errorlevel 1 (
  echo.
  echo Install failed while copying Zotero.dotm.
  pause
  exit /b 1
)

echo.
echo Install finished.
echo Target:
echo %TARGET%
if exist "%BACKUP%" (
  echo Backup:
  echo %BACKUP%
)
pause
