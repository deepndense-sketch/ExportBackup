@echo off
setlocal

set "SRC=D:\Work\Tools\ExportBackup"
set "DEST=%APPDATA%\Adobe\CEP\extensions\ExportBackup"
set "RC=0"

title ExportBackup Deploy

echo Source: %SRC%
echo Destination: %DEST%
echo.

if not exist "%SRC%" (
    echo [ERROR] Source folder not found.
    echo.
    pause
    exit /b 1
)

if not exist "%DEST%" (
    echo Creating CEP extensions folder...
    mkdir "%DEST%"
    if errorlevel 1 (
        echo [ERROR] Could not create destination folder.
        echo.
        pause
        exit /b 1
    )
)

echo Copying ExportBackup files to the CEP extensions folder...
echo.
robocopy "%SRC%" "%DEST%" /MIR /XD .git /XF deploy_extension.bat
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
    echo.
    echo [ERROR] Deployment failed with robocopy exit code %RC%.
    echo The extension may not have been copied correctly.
    echo.
    pause
    exit /b %RC%
)

echo.
echo [DONE] ExportBackup deployed successfully.
echo Restart Premiere Pro if the panel was already open.
echo.
pause
exit /b 0
