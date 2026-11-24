@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0dev_unregister.ps1" %*
