@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0dev_build_reg.ps1" -Activate %*
