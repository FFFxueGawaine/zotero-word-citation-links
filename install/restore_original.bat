@echo off
setlocal
cd /d "%~dp0"

set "BACKUP=%~dp0backup\Zotero.backup.before-linking.dotm"
set "TARGET=%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm"

if not exist "%BACKUP%" (
  echo Backup file not found:
  echo %BACKUP%
  pause
  exit /b 1
)

echo Please close Word first.
copy /y "%BACKUP%" "%TARGET%" >nul

if errorlevel 1 (
  echo Restore failed.
  pause
  exit /b 1
)

echo Restore finished.
pause
