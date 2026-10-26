@echo off
rem ダブルクリックで 検索インデックス作成.ps1 を起動します
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0検索インデックス作成.ps1"
echo.
pause
