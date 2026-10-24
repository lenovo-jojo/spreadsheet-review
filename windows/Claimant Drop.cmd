@echo off
setlocal
if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ClaimantDrop.ps1"
  exit /b
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ClaimantDrop.ps1" %*
