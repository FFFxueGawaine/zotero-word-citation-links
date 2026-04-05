@echo off
setlocal
cd /d "%~dp0"

set "TARGET=%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm"
set "BACKUP=%~dp0backup\Zotero.backup.before-linking.dotm"

echo Please close Microsoft Word before restore.
tasklist /FI "IMAGENAME eq WINWORD.EXE" | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo.
  echo Restore failed: Microsoft Word is still running.
  pause
  exit /b 1
)

if not exist "%BACKUP%" (
  echo.
  echo Restore failed: backup file not found.
  echo Expected:
  echo %BACKUP%
  pause
  exit /b 1
)

copy /y "%BACKUP%" "%TARGET%" >nul
if errorlevel 1 (
  echo.
  echo Restore failed while copying backup Zotero.dotm.
  pause
  exit /b 1
)

echo.
echo Restore finished.
echo Target:
echo %TARGET%
pause
