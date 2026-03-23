@echo off
setlocal enabledelayedexpansion

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
    IF "%PROCESSOR_ARCHITECTURE%" EQU "amd64" (
>nul 2>&1 "%SYSTEMROOT%\SysWOW64\cacls.exe" "%SYSTEMROOT%\SysWOW64\config\system"
) ELSE (
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
)

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params= %*
    echo UAC.ShellExecute "cmd.exe", "/c ""%~s0"" %params:"=""%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------

echo.
echo ================================================================================
echo                    NETWORK ADAPTER DISABLE/ENABLE UTILITY
echo ================================================================================
echo.

REM Open Network Connections control panel
echo Opening Network Connections control panel...
start ncpa.cpl
echo.
echo Please review your network adapters in the Network Connections window.
echo.
pause

echo.
echo ================================================================================
echo                        DETECTED NETWORK ADAPTERS
echo ================================================================================
echo.

REM Get list of network adapters and store them temporarily
set raw_count=0
for /f "skip=3 delims=" %%a in ('netsh interface show interface') do (
    REM Extract adapter name from the line (it's the last column after multiple spaces)
    set "line=%%a"
    REM Remove leading spaces
    for /f "tokens=*" %%b in ("!line!") do set "line=%%b"
    REM Skip header separator line
    echo !line! | findstr "^---" >nul
    if !errorlevel! neq 0 (
        REM Parse the line - format is: State Type Admin Name
        for /f "tokens=1,2,3,*" %%i in ("!line!") do (
            if not "%%l"=="" (
                set "raw_!raw_count!=%%l"
                set /a raw_count+=1
            )
        )
    )
)

REM Sort adapters intelligently (bubble sort with natural number sorting)
set /a last_idx=raw_count-1
for /L %%i in (0,1,%last_idx%) do (
    set /a next_idx=%%i+1
    for /L %%j in (!next_idx!,1,%last_idx%) do (
        call :compare_names "!raw_%%i!" "!raw_%%j!"
        if !compare_result! gtr 0 (
            set "temp=!raw_%%i!"
            set "raw_%%i=!raw_%%j!"
            set "raw_%%j=!temp!"
        )
    )
)

REM Store sorted results
set detected_count=0
for /L %%i in (0,1,%last_idx%) do (
    set "detected_!detected_count!=!raw_%%i!"
    set /a detected_count+=1
)

REM Define letter array
set "letters=ABCDEFGHIJKLMNOPQRSTUVWXYZ"

REM Display lettered list
echo Available network adapters on this system:
echo.
set /a display_count=detected_count-1
for /L %%i in (0,1,%display_count%) do (
    set "letter=!letters:~%%i,1!"
    echo   !letter!. !detected_%%i!
)
echo.
echo ================================================================================
echo.

echo Select adapters to disable/re-enable:
echo   - Enter a LETTER from the list above (A, B, C, etc.)
echo   - Or TYPE the adapter name manually
echo   - Press ENTER without input when you're done
echo.

REM Initialize adapter counter
set adapter_count=0
set adapter_num=1

:input_loop
set "adapter_input="
set /p "adapter_input=Selection %adapter_num% (letter or name, or press Enter to finish): "

if not defined adapter_input goto input_done

REM Check if input is a single letter (selection from list)
set "upper_input=!adapter_input!"
call :to_upper upper_input

REM Find position of letter in alphabet
set letter_pos=-1
for /L %%i in (0,1,25) do (
    set "check_letter=!letters:~%%i,1!"
    if "!upper_input!"=="!check_letter!" set letter_pos=%%i
)

if !letter_pos! geq 0 (
    if !letter_pos! lss !detected_count! (
        REM Valid letter selection
        set "adapter_name=!detected_%letter_pos%!"
        set "adapter_!adapter_count!=!adapter_name!"
        set /a adapter_count+=1
        echo   [OK] Added: !adapter_name!
        set /a adapter_num+=1
        goto input_loop
    ) else (
        echo   [ERROR] Invalid selection. Please choose a letter from A to the last shown letter.
        goto input_loop
    )
)

REM Not a valid letter, treat as adapter name
set "adapter_name=!adapter_input!"

REM Verify adapter exists using netsh
netsh interface show interface name="!adapter_name!" >nul 2>&1
if !errorlevel! equ 0 (
    set "adapter_!adapter_count!=!adapter_name!"
    set /a adapter_count+=1
    echo   [OK] Added: !adapter_name!
    set /a adapter_num+=1
    goto input_loop
) else (
    echo   [ERROR] Adapter '!adapter_name!' not found. Please check the name and try again.
    goto input_loop
)

:input_done
if %adapter_count% equ 0 (
    echo.
    echo No adapters specified. Exiting.
    echo.
    pause
    exit /b
)

echo.
echo You have selected %adapter_count% adapter(s) to disable/re-enable:
set /a last_index=%adapter_count%-1
for /L %%i in (0,1,%last_index%) do (
    if defined adapter_%%i echo   - !adapter_%%i!
)
echo.

REM Ask for disable duration
set disable_time=30

:time_input
set "user_time="
set /p "user_time=Enter disable duration in seconds (1-300, press Enter for default 30): "

if not defined user_time (
    echo.
    echo Using default: 30 seconds
    goto time_done
)

REM Validate it's a number
set /a test_num=user_time 2>nul
if !test_num! equ 0 (
    if not "!user_time!"=="0" (
        echo   [ERROR] Invalid input. Please enter a number.
        goto time_input
    )
)

REM Check if value is too high
if !user_time! gtr 300 (
    echo   [ERROR] Value too high. Maximum is 300 seconds.
    goto time_input
)

REM Check if value is too low (0 or negative)
if !user_time! lss 1 (
    echo   [WARNING] Value too low. Setting to minimum 30 seconds.
    set disable_time=30
    goto time_done
)

REM Valid input
set disable_time=!user_time!

:time_done
echo.
echo Disable duration set to: %disable_time% seconds
echo.

REM Additional confirmation for long durations
if %disable_time% gtr 120 (
    echo WARNING: You have selected a duration greater than 120 seconds ^(%disable_time%s^).
    set /p "long_confirm=Are you sure you want to proceed? (Y/N): "
    if /i not "!long_confirm!"=="Y" (
        echo.
        echo Operation cancelled.
        echo.
        pause
        exit /b
    )
    echo.
)

set /p "confirm=Proceed with disabling these adapters for %disable_time% seconds? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo.
    echo Operation cancelled.
    echo.
    pause
    exit /b
)

echo.
echo ================================================================================
echo                            DISABLING ADAPTERS
echo ================================================================================
echo.

REM Disable all adapters
set /a last_index=%adapter_count%-1
for /L %%i in (0,1,%last_index%) do (
    if defined adapter_%%i (
        echo Disabling adapter: !adapter_%%i!
        netsh interface set interface name="!adapter_%%i!" admin=disable
        if !errorlevel! equ 0 (
            echo   [OK] Disabled successfully
        ) else (
            echo   [ERROR] Failed to disable !adapter_%%i!
        )
        echo.
    )
)

echo ================================================================================
echo Waiting for %disable_time% seconds...
echo ================================================================================
echo.

REM Countdown timer
for /L %%i in (%disable_time%,-1,1) do (
    <nul set /p "=Time remaining: %%i seconds   "
    timeout /t 1 >nul
    echo.
)

echo.
echo ================================================================================
echo                            RE-ENABLING ADAPTERS
echo ================================================================================
echo.

REM Re-enable all adapters
set /a last_index=%adapter_count%-1
for /L %%i in (0,1,%last_index%) do (
    if defined adapter_%%i (
        echo Re-enabling adapter: !adapter_%%i!
        netsh interface set interface name="!adapter_%%i!" admin=enable
        if !errorlevel! equ 0 (
            echo   [OK] Enabled successfully
        ) else (
            echo   [ERROR] Failed to enable !adapter_%%i!
        )
        echo.
    )
)

echo.
echo ================================================================================
echo                      OPERATION COMPLETED SUCCESSFULLY!
echo ================================================================================
echo.
pause
goto :eof

REM ===== Subroutines =====

:compare_names
REM Compare two adapter names for natural sorting
REM Returns: compare_result = -1 (less), 0 (equal), 1 (greater)
set "name1=%~1"
set "name2=%~2"

REM Try to find last space followed by a number
call :extract_base_and_num "%name1%" base1 num1
call :extract_base_and_num "%name2%" base2 num2

REM If base names are different, do string comparison
if not "%base1%"=="%base2%" (
    if "%name1%" LSS "%name2%" (
        set compare_result=-1
    ) else if "%name1%" GTR "%name2%" (
        set compare_result=1
    ) else (
        set compare_result=0
    )
    goto :eof
)

REM Base names are same, compare numbers
if %num1% LSS %num2% (
    set compare_result=-1
) else if %num1% GTR %num2% (
    set compare_result=1
) else (
    set compare_result=0
)
goto :eof

:extract_base_and_num
REM Extract base name and trailing number from adapter name
REM Usage: call :extract_base_and_num "adapter name" var_base var_num
set "full_name=%~1"
set "result_base=%full_name%"
set "result_num=0"

REM Check if name ends with a space followed by digits
for /f "tokens=1* delims= " %%a in ("%full_name%") do (
    set "first_part=%%a"
    set "rest_part=%%b"
)

REM If there's a second part, check if it's a number
if defined rest_part (
    set /a test_num=rest_part 2>nul
    if !test_num! gtr 0 (
        REM It's a number, so split it
        set "result_base=%first_part%"
        set "result_num=!rest_part!"
    ) else (
        REM Not a number, check if rest has more parts
        for /f "tokens=1* delims= " %%c in ("!rest_part!") do (
            set "second_part=%%c"
            set "third_part=%%d"
        )
        if defined third_part (
            set /a test_num2=third_part 2>nul
            if !test_num2! gtr 0 (
                set "result_base=%first_part% !second_part!"
                set "result_num=!third_part!"
            )
        )
    )
)

set "%~2=%result_base%"
set "%~3=%result_num%"
goto :eof

:to_upper
REM Convert string to uppercase
set "str=!%~1!"
for %%a in (a b c d e f g h i j k l m n o p q r s t u v w x y z) do (
    call set "str=%%str:%%a=%%a%%"
)
for %%a in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    call set "str=%%str:%%a=%%a%%"
)
set "str=%str:a=A%"
set "str=%str:b=B%"
set "str=%str:c=C%"
set "str=%str:d=D%"
set "str=%str:e=E%"
set "str=%str:f=F%"
set "str=%str:g=G%"
set "str=%str:h=H%"
set "str=%str:i=I%"
set "str=%str:j=J%"
set "str=%str:k=K%"
set "str=%str:l=L%"
set "str=%str:m=M%"
set "str=%str:n=N%"
set "str=%str:o=O%"
set "str=%str:p=P%"
set "str=%str:q=Q%"
set "str=%str:r=R%"
set "str=%str:s=S%"
set "str=%str:t=T%"
set "str=%str:u=U%"
set "str=%str:v=V%"
set "str=%str:w=W%"
set "str=%str:x=X%"
set "str=%str:y=Y%"
set "str=%str:z=Z%"
set "%~1=%str%"
goto :eof
