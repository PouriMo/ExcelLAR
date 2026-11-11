@echo off
rem Installs the ExcelLAR add-in for the current user. See INSTALL.md.
rem PowerShell runs this one script without changing any system setting.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install\Install-ExcelLAR.ps1" %*
echo.
pause
