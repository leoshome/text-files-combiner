@echo off
cd /d "%~dp0"
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ".\text-files-combiner.ps1"