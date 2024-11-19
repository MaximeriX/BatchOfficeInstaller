@echo off
@setlocal DisableDelayedExpansion

echo ::::::::::::::::::::::::::::::::: 
echo ADMINISTRATOR PRIVILEGES REQUIRED
echo :::::::::::::::::::::::::::::::::

@rem Checking for administrator privileges
net session
if %errorlevel% neq 0 goto runadmin
goto adminstart

@rem Running cmd as administrator
:runadmin
CD /d %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
timeout 3 && exit

@rem Script start
:adminstart
:start
cls
@rem Creating OfficeSetup folder on C:\
mkdir C:\OfficeSetup
cls
echo Have you moved the program folder to C:\?
echo [1] Yes
echo [2] Move
choice /C:12 /M "Enter your choice:" /N
set ChoiceFiles=%errorlevel%
if %ChoiceFiles% == 1 goto bitselect
if %ChoiceFiles% == 2 goto movefiles

:movefiles
cls
echo Copying the program to C:\...
xcopy "%~dp0*" C:\OfficeSetup /E /I
cls
cd C:\OfficeSetup
start C:\OfficeSetup\Setup.bat
exit

:bitselect
cls
echo Select system type (Start - Settings - System - About - System type)
echo [1] 64-Bit
echo [2] 32-Bit
choice /C:12 /M "Enter your choice:" /N
set ChoiceBit=%errorlevel%
if %ChoiceBit% == 1 goto warn64
if %ChoiceBit% == 2 goto warn32

@rem Warning screen for 32-bit system
:warn32
cls
echo Previous versions of Office will be removed! Continue?
echo [1] Yes
echo [2] No
choice /C:12 /M "Enter your choice:" /N
set ChoiceWarn=%errorlevel%
if %ChoiceWarn% == 1 goto office32
if %ChoiceWarn% == 2 goto exit

@rem Warning screen for 64-bit system
:warn64
cls
echo Previous versions of Office will be removed! Continue?
echo [1] Yes
echo [2] No
choice /C:12 /M "Enter your choice:" /N
set ChoiceWarn=%errorlevel%
if %ChoiceWarn% == 1 goto office64
if %ChoiceWarn% == 2 goto exit

@rem Loading for 64-bit system
:office64
cls
echo Select the version of Office you want to install
echo [1] Office LTSC 2024
echo [2] Office LTSC 2021
echo [3] Office LTSC 2019
choice /C:123 /M "Enter your choice:" /N
set ChoiceOffice=%errorlevel%
if %ChoiceOffice% == 1 goto OfficeLTSC2464
if %ChoiceOffice% == 2 goto OfficeLTSC2164
if %ChoiceOffice% == 3 goto OfficeLTSC1964

@rem Loading for 32-bit system
:office32
cls
echo Select the version of Office you want to install
echo [1] Office LTSC 2024
echo [2] Office LTSC 2021
echo [3] Office LTSC 2019
choice /C:123 /M "Enter your choice:" /N
set ChoiceOffice=%errorlevel%
if %ChoiceOffice% == 1 goto OfficeLTSC32432
if %ChoiceOffice% == 2 goto OfficeLTSC32132
if %ChoiceOffice% == 3 goto OfficeLTSC31932

@rem Office 32-bit and 64-bit setup
:OfficeLTSC32432
cd C:\OfficeSetup\OfficeInstaller
setlocal
@rem URL to download
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Destination folder
set DEST=C:\OfficeSetup\OfficeInstaller
@rem File name
set FILE=OfficeExtracter.exe
@rem Ensure the destination folder exists
if not exist %DEST% (
    mkdir %DEST%
)
@rem Downloading the file using curl
echo Downloading the file...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%\%FILE%
) else (
    echo Error downloading file.
)
endlocal 
cls
start C:\OfficeSetup\OfficeInstaller\OfficeExtracter.exe /extract:C:\OfficeSetup\OfficeInstaller /passive /norestart /quiet
echo Please wait 10 seconds.
timeout 10 && echo Loading Office LTSC 2024 32-Bit
start C:\OfficeSetup\OfficeInstaller\setup.exe /configure config24-32.xml 
timeout 10 && exit

:OfficeLTSC2464
cd C:\OfficeSetup\OfficeInstaller
setlocal
@rem URL to download
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Destination folder
set DEST=C:\OfficeSetup\OfficeInstaller
@rem File name
set FILE=OfficeExtracter.exe
@rem Ensure the destination folder exists
if not exist %DEST% (
    mkdir %DEST%
)
@rem Downloading the file using curl
echo Downloading the file...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%\%FILE%
) else (
    echo Error downloading file.
)
endlocal 
cls
start C:\OfficeSetup\OfficeInstaller\OfficeExtracter.exe /extract:C:\OfficeSetup\OfficeInstaller /passive /norestart /quiet
echo Please wait 10 seconds.
timeout 10 && echo Loading Office LTSC 2024 64-Bit
start C:\OfficeSetup\OfficeInstaller\setup.exe /configure config24-64.xml 
timeout 10 && exit

:OfficeLTSC32132
cd C:\OfficeSetup\OfficeInstaller
setlocal
@rem URL to download
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Destination folder
set DEST=C:\OfficeSetup\OfficeInstaller
@rem File name
set FILE=OfficeExtracter.exe
@rem Ensure the destination folder exists
if not exist %DEST% (
    mkdir %DEST%
)
@rem Downloading the file using curl
echo Downloading the file...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%\%FILE%
) else (
    echo Error downloading file.
)
endlocal 
cls
start C:\OfficeSetup\OfficeInstaller\OfficeExtracter.exe /extract:C:\OfficeSetup\OfficeInstaller /passive /norestart /quiet
echo Please wait 10 seconds.
timeout 10 && echo Loading Office LTSC 2021 32-Bit
start C:\OfficeSetup\OfficeInstaller\setup.exe /configure config21-32.xml 
timeout 10 && exit

:OfficeLTSC2164
cd C:\OfficeSetup\OfficeInstaller
setlocal
@rem URL to download
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Destination folder
set DEST=C:\OfficeSetup\OfficeInstaller
@rem File name
set FILE=OfficeExtracter.exe
@rem Ensure the destination folder exists
if not exist %DEST% (
    mkdir %DEST%
)
@rem Downloading the file using curl
echo Downloading the file...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%\%FILE%
) else (
    echo Error downloading file.
)
endlocal
cls
start C:\OfficeSetup\OfficeInstaller\OfficeExtracter.exe /extract:C:\OfficeSetup\OfficeInstaller /passive /norestart /quiet
echo Please wait 10 seconds.
timeout 10 && echo Installing Office LTSC 2021 64-Bit
start C:\OfficeSetup\OfficeInstaller\setup.exe /configure config21-64.xml 
timeout 10 && exit

:OfficeLTSC31932
cd C:\OfficeSetup\OfficeInstaller
setlocal
@rem URL to download
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Destination folder
set DEST=C:\OfficeSetup\OfficeInstaller
@rem File name
set FILE=OfficeExtracter.exe
@rem Ensure the destination folder exists
if not exist %DEST% (
    mkdir %DEST%
)
@rem Downloading the file using curl
echo Downloading the file...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%\%FILE%
) else (
    echo Error downloading file.
)
endlocal 
cls
start C:\OfficeSetup\OfficeInstaller\OfficeExtracter.exe /extract:C:\OfficeSetup\OfficeInstaller /passive /norestart /quiet
echo Please wait 10 seconds.
timeout 10 && echo Installing Office LTSC 2019 32-Bit
start C:\OfficeSetup\OfficeInstaller\setup.exe /configure config19-32.xml 
timeout 10 && exit

:OfficeLTSC1964
cd C:\OfficeSetup\OfficeInstaller
setlocal
@rem URL to download
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Destination folder
set DEST=C:\OfficeSetup\OfficeInstaller
@rem File name
set FILE=OfficeExtracter.exe
@rem Ensure the destination folder exists
if not exist %DEST% (
    mkdir %DEST%
)
@rem Downloading the file using curl
echo Downloading the file...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%\%FILE%
) else (
    echo Error downloading file.
)
endlocal 
cls
start C:\OfficeSetup\OfficeInstaller\OfficeExtracter.exe /extract:C:\OfficeSetup\OfficeInstaller /passive /norestart /quiet
echo Please wait 10 seconds.
timeout 10 && echo Installing Office LTSC 2019 64-Bit
start C:\OfficeSetup\OfficeInstaller\setup.exe /configure config19-64.xml 
timeout 10 && exit

:exit
exit