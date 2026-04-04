@echo off
setlocal
cd /d "%~dp0"

echo Please close Microsoft Word before install.
echo.
echo [1/1] Running installer...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_zotero_word_links.ps1"

if errorlevel 1 (
  echo.
  echo Install failed.
  pause
  exit /b 1
)

echo.
echo Install finished.
pause
