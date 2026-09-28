@echo off
title Google Takeout Metadata Fixer v6.2.3
setlocal
set "GUI=%~dp0GoogleTakeoutMetadataFixer-GUI.ps1"

if not exist "%GUI%" (
  echo Could not find GoogleTakeoutMetadataFixer-GUI.ps1 next to this launcher.
  pause
  exit /b 1
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%GUI%"
if errorlevel 1 (
  echo.
  echo The fixer closed with an error. Review any message shown above.
  pause
)
