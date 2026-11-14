@echo off
setlocal
cd /d "%~dp0"
set "UI=%~dp0Scripts\Start-PackageUI.ps1"
echo Starting Application Packaging Automation - V2
echo Script: "%UI%"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%UI%"
exit /b %ERRORLEVEL%
