@echo off
REM Bootstrapper to launch the IVDR validation GUI without requiring admin rights.
REM Assumes this batch file resides in the same directory as Main.ps1.

setlocal
set SCRIPT_DIR=%~dp0

REM Use -ExecutionPolicy Bypass to allow loading unsigned scripts in a controlled
REM environment.  -NoProfile ensures a clean run.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Main.ps1"

endlocal

