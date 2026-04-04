@echo off
setlocal
cd /d "%~dp0"

echo Please close Microsoft Word before install.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_zotero_word_links.ps1"

if errorlevel 1 (
  echo.
  echo Install failed.
  if /i not "%ZWL_NO_PAUSE%"=="1" pause
  exit /b 1
)

echo.
echo Install finished.
if /i not "%ZWL_NO_PAUSE%"=="1" pause
