@echo off
rem Double-click to install the ProgressLogger macro (creates the .xlsm next to the .xlsx).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-ProgressLogger.ps1" -EnableVbomAccess
echo.
pause
