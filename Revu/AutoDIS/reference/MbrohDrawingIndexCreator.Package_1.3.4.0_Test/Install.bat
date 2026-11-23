:: Self-elevate if not already admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    setlocal enabledelayedexpansion
    set "TARGET=%~f0"
    set "WORKDIR=%~dp0"
    set "TARGET=!TARGET:P:=\\fs01\public!"
    set "WORKDIR=!WORKDIR:P:=\\fs01\public!"
::powershell -Command "Start-Process -FilePath '!TARGET!' -Verb RunAs -WorkingDirectory '!WORKDIR!'"
powershell -Command "Start-Process -FilePath '!TARGET!' -Verb Run -WorkingDirectory '!WORKDIR!'"
    endlocal
    exit /b
)
net use P: \\fs01\public
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Install.ps1"
::pause