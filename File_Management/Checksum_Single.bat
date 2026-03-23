@echo off
echo ====================================
echo    SHA256 Hash Calculator
echo ====================================
echo.
echo Paste a file path, or press Enter to browse:
echo.
set /p input="File path: "

if "%input%"=="" goto browse

:: User entered a path
set filepath=%input:"=%

:: Check if file exists
if not exist "%filepath%" (
    echo Error: File not found.
    pause
    exit /b
)
goto calculate

:browse
:: Use PowerShell to open a file selection dialog
for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Title = 'Select a file to calculate SHA256'; $ofd.Filter = 'All Files (*.*)|*.*'; if ($ofd.ShowDialog() -eq 'OK') { $ofd.FileName }"`) do set filepath=%%i

:: Check if user canceled
if "%filepath%"=="" (
    echo No file selected. Exiting.
    pause
    exit /b
)
goto calculate

:manual
echo.
set /p filepath="Enter the full file path: "
if "%filepath%"=="" (
    echo No file path entered. Exiting.
    pause
    exit /b
)
:: Remove quotes if user added them
set filepath=%filepath:"=%

:: Check if file exists
if not exist "%filepath%" (
    echo Error: File not found.
    pause
    exit /b
)
goto calculate

:calculate
:: Compute SHA256 hash
echo.
echo Calculating SHA256 hash for: %filepath%
echo.
certutil -hashfile "%filepath%" SHA256

pause