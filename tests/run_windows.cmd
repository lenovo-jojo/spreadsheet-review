@echo off
rem Runs the Windows tool without dialogs so tests/test_workbook.py can check it.
rem Usage: set CLAIMANTS_BINARY=%CD%\tests\run_windows.cmd ^&^& python -m unittest discover -s tests -v
setlocal
set CLAIMANTS_NO_DIALOG=1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\windows\ClaimantDrop.ps1" %*
