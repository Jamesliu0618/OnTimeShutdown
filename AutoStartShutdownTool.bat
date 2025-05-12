@echo off
:: Auto start script for Shutdown Tool
:: Place this script in Windows Startup folder

:: Navigate to the application directory
cd /d %~dp0

:: Start the application minimized
start "" "%~dp0bin\Debug\net8.0\Showdown.exe" --minimized

exit 