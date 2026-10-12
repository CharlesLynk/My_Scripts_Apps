@echo off
:: Set the path to accoreconsole.exe
set accoreexe="C:\Program Files\Autodesk\AutoCAD 2022\accoreconsole.exe"

:: Set the path to the folder containing the drawings
set source="C:\pw_pl1\LYNKCN\d4544806\empty"

:: Set the path to the script file
set script="C:\Users\lynkcn\Documents\AutoCAD Macros_Lisp_VBA\AutoCAD Scripts\ACCoreConsole\TitleBlockChange_Empty.scr"

:: Loop through each DWG file in the folder and run the script
for %%f in ("%source%\*.dwg") do (
    %accoreexe% /i "%%f" /s %script%
)

pause
