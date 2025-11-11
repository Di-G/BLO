@echo off
setlocal
set PORT=5500
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0serve.ps1" -Port %PORT%
endlocal


