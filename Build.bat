@echo off
REM Quick build script - roept PowerShell script aan
powershell -ExecutionPolicy Bypass -File "%~dp0Build-VBEAddIn.ps1" %*
pause
