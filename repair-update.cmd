@echo off
setlocal
title Backup Project Update Repair
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0repair-update.ps1"
echo.
pause
