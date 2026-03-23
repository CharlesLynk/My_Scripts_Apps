@echo off
setlocal enabledelayedexpansion

:: Check if "checksums_old.txt" exists and delete it
if exist "checksums_old.txt" (
    del "checksums_old.txt"
)

:: Check if "checksums" file exists
if exist "checksums.txt" (
    :: Rename existing "checksums" to "checksums_old"
    ren "checksums.txt" "checksums_old.txt"
)

:: Generate MD5 checksum for each file in the current directory
for %%f in (*) do (
    :: Exclude "checksums.bat", "checksums", and "checksums_old" from the loop
    if /i "%%~nxf" neq "checksum.bat" (
        if /i "%%~nxf" neq "checksums.txt" (
            if /i "%%~nxf" neq "checksums_old.txt" (
                :: Append the filename next to the checksum
                echo %%~nxf >> "checksums.txt"
                :: Calculate MD5 hash for the file
                certutil -hashfile "%%f" MD5 | find /i /v "md5" | find /i /v "CertUtil" >> "checksums.txt"
				echo. >> "checksums.txt"
            )
        )
    )
)

:: Display completion message
echo Checksums saved to "checksums" file.
endlocal