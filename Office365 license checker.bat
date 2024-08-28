@echo off
title Office 365 Checker by AykhanUV
color 3
cls
echo Checking Office 365 Information...
echo.
echo 1. Check License Status
echo 2. Check Activation Expiry
echo.

:: Use choice command to get user input
choice /C 12 /N /M "Enter your choice (1 or 2): "

:: Directly execute the selected option
if errorlevel 2 goto Expiry
if errorlevel 1 goto LicenseStatus

:LicenseStatus
cls
echo Checking license status...
cscript /nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus | findstr /R /C:"LICENSE NAME:" /C:"LICENSE STATUS:" /C:"Last 5 characters of installed product key:"
pause
goto end

:Expiry
cls
echo Checking activation expiry...
cscript /nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus | findstr /R /C:"REMAINING GRACE:" | powershell -Command "$input -replace 'REMAINING GRACE:', 'Activation Expiry:'"
pause
goto end

:end
exit
