@echo off
setlocal

set "SCRIPT=%~dp0Show-DiskUsage.ps1"
set "POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%SCRIPT%" (
    echo ERRO: Nao encontrei "%SCRIPT%".
    pause
    exit /b 1
)

net session >nul 2>&1
if not "%errorlevel%"=="0" (
    "%POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%POWERSHELL%' -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%""'"
    exit /b 0
)

"%POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
exit /b %errorlevel%
