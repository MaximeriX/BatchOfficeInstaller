@chcp 65001
@echo off
@setlocal DisableDelayedExpansion
title ADMINISTRATOR PRIVILEGES REQUIRED
@rem Checking for administrator privileges
net session
cls
echo.
echo    ╠═══════════════════════════════════╣
echo      ADMINISTRATOR PRIVILEGES REQUIRED  
echo    ╠═══════════════════════════════════╣
echo.
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
title Office Installer by MaximeriX
cls
echo Office includes PowerPoint, Word, Excel, Outlook, and OneNote.
echo ╔═╦══════╗
echo ║1║ Okay ║
echo ║2║ Exit ║
echo ╚═╩═══╦══╝
choice /C:12 /M "Enter ╚→ :" /N
set ChoiceOk=%errorlevel%
if %ChoiceOk% == 1 goto bitselect
if %ChoiceOk% == 2 goto exit

@rem Selecting the system type
:bitselect
cls
echo Select the system type (Start → Settings → System → About → System type)
echo ╔═╦════════╗
echo ║1║ 64-Bit ║
echo ║2║ 32-Bit ║
echo ╚═╩═══╦════╝
choice /C:12 /M "Enter ╚→ :" /N
set ChoiceBit=%errorlevel%
if %ChoiceBit% == 1 (set OCE=64) && goto office
if %ChoiceBit% == 2 (set OCE=32) && goto office

@rem Setting up the download
:office
mkdir %~dp0OfficeSetupFiles
cls
echo Select the version of Office you want to install
echo ╔═╦══════════════════╗
echo ║1║ Office LTSC 2024 ║
echo ║2║ Office LTSC 2021 ║
echo ║3║ Office 2019      ║
echo ╚═╩═══╦══════════════╝
choice /C:123 /M "Enter ╚→ :" /N
set ChoiceOffice=%errorlevel%
if %ChoiceOffice% == 1 goto LTSC2024
if %ChoiceOffice% == 2 goto LTSC2021
if %ChoiceOffice% == 3 goto Office2019

@rem Settings for Office LTSC 2024
:LTSC2024
set CID=ef5c8a1f-1356-46fc-984b-634b44e23987
set Channel=PerpetualVL2024
set PID=ProPlus2024Volume
set PIDKEY=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
set OfficeVer=Office LTSC 2024
goto ConfigGen2124

@rem Settings for Office LTSC 2021
:LTSC2021
set CID=c04f0bb9-2868-4356-8632-88c4c1a4870c
set Channel=PerpetualVL2021
set PID=ProPlus2024Volume
set PIDKEY=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
set OfficeVer=Office LTSC 2021
goto ConfigGen2124

@rem Settings for Office 2019
:Office2019
set CID=906df582-99a6-4c42-95e0-a13f220cd505
set Channel=PerpetualVL2019
set PID=ProPlus2019Volume
set PIDKEY=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
set OfficeVer=Office 2019
goto ConfigGen19

@rem Generating XML file for Office LTSC 2024 and LTSC 2021
:ConfigGen2124
cls
echo Generating configuration XML file...
(
    echo ^<Configuration ID="%CID%"^>
    echo   ^<Add OfficeClientEdition="%OCE%" Channel="%Channel%"^>
    echo     ^<Product ID="%PID%" PIDKEY="%PIDKEY%"^>
    echo       ^<Language ID="en-us" /^>
    echo       ^<ExcludeApp ID="Access" /^>
    echo       ^<ExcludeApp ID="Lync" /^>
    echo       ^<ExcludeApp ID="OneDrive" /^>
    echo       ^<ExcludeApp ID="Publisher" /^>
    echo     ^</Product^>
    echo   ^</Add^>
    echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>
    echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /^>
    echo   ^<Property Name="DeviceBasedLicensing" Value="0" /^>
    echo   ^<Property Name="SCLCacheOverride" Value="0" /^>
    echo   <Property Name="AUTOACTIVATE" Value="1" /^>
    echo   ^<Updates Enabled="TRUE" /^>
    echo   ^<AppSettings^>
    echo     ^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^>
    echo     ^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^>
    echo     ^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^>
    echo   ^</AppSettings^>
    echo   ^<Display Level="Full" AcceptEULA="TRUE" /^>
    echo ^</Configuration^>
) > %~dp0OfficeSetupFiles\Config.xml
timeout 2
goto OfficeSetup

@rem Generating XML file for Office 2019
:ConfigGen19
cls
echo Generating configuration XML file...
(
    echo ^<Configuration ID="%CID%"^>
    echo   ^<Add OfficeClientEdition="%OCE%" Channel="%Channel%"^>
    echo     ^<Product ID="%PID%" PIDKEY="%PIDKEY%"^>
    echo       ^<Language ID="en-us" /^>
    echo       ^<ExcludeApp ID="Access" /^>
    echo       ^<ExcludeApp ID="Groove" /^>
    echo       ^<ExcludeApp ID="Lync" /^>
    echo       ^<ExcludeApp ID="OneDrive" /^>
    echo       ^<ExcludeApp ID="Publisher" /^>
    echo     ^</Product^>
    echo   ^</Add^>
    echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>
    echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /^>
    echo   ^<Property Name="DeviceBasedLicensing" Value="0" /^>
    echo   ^<Property Name="SCLCacheOverride" Value="0" /^>
    echo   <Property Name="AUTOACTIVATE" Value="1" /^>
    echo   ^<Updates Enabled="TRUE" /^>
    echo   ^<AppSettings^>
    echo     ^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^>
    echo     ^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^>
    echo     ^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^>
    echo   ^</AppSettings^>
    echo   ^<Display Level="Full" AcceptEULA="TRUE" /^>
    echo ^</Configuration^>
) > %~dp0OfficeSetupFiles\Config.xml
timeout 2
goto OfficeSetup

@rem Downloading Office
:OfficeSetup
setlocal
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
set DEST=%~dp0OfficeSetupFiles\
set FILE=OfficeExtracter.exe
@rem Downloading the file using curl
echo Downloading the file %FILE%...
curl -L -o %DEST%\%FILE% %URL%
@rem Checking if the file was successfully downloaded
if exist %DEST%\%FILE% (
    echo File successfully downloaded to %DEST%%FILE%
) else (
    echo File download error.
)
endlocal
timeout 2
cls
start %~dp0OfficeSetupFiles\OfficeExtracter.exe /extract:%~dp0OfficeSetupFiles\ /passive /norestart /quiet
echo Please wait 5 seconds...
timeout 5 && del /f %~dp0OfficeSetupFiles\configuration-Office365-x64.xml && cls && echo Downloading %OfficeVer% %OCE%-Bits...
start %~dp0OfficeSetupFiles\setup.exe /configure Config.xml
timeout 10 && exit

:exit
exit