@echo off
rem Removes the ExcelLAR add-in for the current user. OpenSolver is not changed.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install\Uninstall-ExcelLAR.ps1" %*
echo.
pause
