@echo off
setlocal

set "SRC=D:\Work\Tools\ExportBackup"
set "DEST=C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\ExportBackup"

echo Source: %SRC%
echo Destination: %DEST%
echo.

if not exist "%SRC%" (
    echo Source folder not found.
    exit /b 1
)

if not exist "%DEST%" (
    mkdir "%DEST%"
)

robocopy "%SRC%" "%DEST%" /MIR /XD .git /XF deploy_extension.bat
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
    echo Deployment failed with robocopy exit code %RC%.
    exit /b %RC%
)

echo.
echo ExportBackup deployed successfully.
echo Restart Premiere Pro if the panel was already open.
exit /b 0
