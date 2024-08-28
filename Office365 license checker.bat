@echo off
title Made by AykhanUV
color 3
cls

echo Checking for Office 365 installation...
echo.

:: Check if Office 365 is installed
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    echo Office 365 found!
    echo.
    echo 1. Check License Status
    echo 2. Check Activation Expiry
    echo.

    :: Use choice command to get user input
    choice /C 12 /N /M "Enter your choice (1 or 2): "

    :: Directly execute the selected option
    if errorlevel 2 goto Expiry
    if errorlevel 1 goto LicenseStatus
) else (
    echo Office 365 is not installed on this system.
    pause
    goto end
)

:LicenseStatus
cls
echo Checking license status...
cscript /nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus | findstr /R /C:"LICENSE NAME:" /C:"LICENSE STATUS:" /C:"Last 5 characters of installed product key:" || echo No valid license found.
pause
goto end

:Expiry
cls
echo Checking activation expiry...
cscript /nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus | findstr /R /C:"REMAINING GRACE:" || echo No valid license found.
pause
goto end

:end
exit
