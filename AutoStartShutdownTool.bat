@echo off
:: 自動關機工具的自動啟動腳本
:: 將此腳本放入 Windows 啟動資料夾

:: 導航到應用程式目錄
cd /d %~dp0

:: 以最小化狀態啟動應用程式
start "" "%~dp0bin\Debug\net8.0-windows\Showdown.exe" --minimized

exit 