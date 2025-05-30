@echo off
setlocal EnableDelayedExpansion

set found=0
set "files="

:: Look for TestLab files in the current folder
for %%F in (*.*) do (
    set "name=%%~nF"
    set "ext=%%~xF"
    call :checkFile "%%F" "!name!" "!ext!"
)

:: Check if exactly 4 files found, otherwise exit with error
if !found! NEQ 4 (
    echo.
    echo ,-------------------,     ,--------,
    echo ^| ztl - Zip TestLab ^| --^> ^| FAILED ^|
    echo +-------------------'     '--------'
    echo ztl needs 4 files, but there are !found! matching files.
    echo Please, prepare them correctly and try again.
    echo.& echo.
    pause
    goto :eof
)

:: Create archive name
set "archive=%prefix%.7z"

echo ,-------------------------------------------------------------------------,
echo ^| ztl - Zip TestLab - Launching creation of %archive% using 7-zip: ^|
echo 'v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v'
echo.

:: Call 7z with all files in %files%
for /f "delims=" %%L in ('^
    7z a -bb1 -bd -y "%archive%" %files% 2^>nul ^| ^
    findstr /R "[0-9][0-9][0-9][0-9]_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" ^
') do (
    echo %%L
)

echo.
goto :eof

:: ---------------------------------------------------------
:checkFile
setlocal EnableDelayedExpansion
set "full=%~1"
set "name=%~2"
set "ext=%~3"

:: Check extension
if /I not "!ext!"==".pdf" if /I not "!ext!"==".txt" if /I not "!ext!"==".html" (
    endlocal
    goto :eof
)

:: Check filename
echo !name! | findstr /I "DiagAsFound DiagAsLeft CalAsFound CalAsLeft" >nul
if errorlevel 1 (
    endlocal
    goto :eof
)

:: Check YEAR_SERIALNO pattern (4 digits underscore 8 digits somewhere)
echo !name! | findstr /R "^[0-9][0-9][0-9][0-9]_.*[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" >nul
if errorlevel 1 (
    endlocal
    goto :eof
)

:: Valid file found - commit variables to outer environment
endlocal & (
    set /a found+=1
    set "file%found%=%full%"
    if defined files (
        set "files=!files! "%full%""
    ) else (
        set "files="%full%""
    )
)

:: Extract prefix from first file only
if %found%==1 (
    for /F "tokens=1,2 delims=_" %%A in ("%name%") do (
        set "prefix=%%A_%%B"
    )
)

goto :eof
