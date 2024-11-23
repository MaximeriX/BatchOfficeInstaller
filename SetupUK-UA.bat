@chcp 65001
@echo off
@setlocal EnableDelayedExpansion
title ПОТРІБНІ ПРИВІЛЕЇ АДМІНІСТРАТОРА

@rem Перевірка привілеїв адміністратора
net session
cls
echo.
echo    ╠══════════════════════════════════╣
echo      ПОТРІБНІ ПРИВІЛЕЇ АДМІНІСТРАТОРА  
echo    ╠══════════════════════════════════╣
echo.
if %errorlevel% neq 0 goto runadmin
goto adminstart

@rem Запуск cmd як адміністратор
:runadmin
CD /d %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
timeout 3 && exit

@rem Початок скрипта
:adminstart
:start
title Office Installer by MaximeriX
cls
echo Office - має в собі Access, Excel, OneNote, Outlook, PowerPoint, Publisher, Word. (Можна змінити)
echo ╔═╦═══════╗
echo ║1║ Добре ║
echo ║2║ Вийти ║
echo ╚═╩═════╦═╝
choice /C:12 /M "Введіть ╚→ :" /N
set ChoiceOk=%errorlevel%
if %ChoiceOk% == 1 goto ExcludeApps
if %ChoiceOk% == 2 goto exit

@rem Вибір додатків для виключення
:ExcludeApps
cls
echo Виберіть додатки, які не хочете завантажувати.
echo ╔═╦════════════╗
echo ║1║ Access     ║
echo ║2║ Excel      ║
echo ║3║ OneNote    ║
echo ║4║ Outlook    ║
echo ║5║ PowerPoint ║
echo ║6║ Publisher  ║
echo ║7║ Word       ║
echo ╚═╩════════════╩═══╗
set "exclude="
set /p input="Введіть (1 3 6 тд) ╚→ : "

for %%i in (%input%) do (
    if %%i==1 (
        set Access=Access
    ) else if %%i==2 (
        set Excel=Excel
    ) else if %%i==3 (
        set OneNote=OneNote
    ) else if %%i==4 (
        set Outlook=Outlook
    ) else if %%i==5 (
        set PowerPoint=PowerPoint
    ) else if %%i==6 (
        set Publisher=Publisher
    ) else if %%i==7 (
        set Word=Word
    ) else (
        echo Недійсний вибір: %%i
        timeout 2 && exit
    )
)
goto bitselect

@rem Вибір типу системи
:bitselect
cls
echo Виберіть тип системи (Пуск → Налаштування → Система → Про → Тип системи)
echo ╔═╦══════════╗
echo ║1║ 64-Бітна ║
echo ║2║ 32-Бітна ║
echo ╚═╩═════╦════╝
choice /C:12 /M "Введіть ╚→ :" /N
set ChoiceBit=%errorlevel%
if %ChoiceBit% == 1 (set OCE=64) && goto office
if %ChoiceBit% == 2 (set OCE=32) && goto office

@rem Налаштування завантажування
:office
mkdir %~dp0OfficeSetupFiles
cls
echo Виберіть версію Office, яку хочете встановити
echo ╔═╦══════════════════╗
echo ║1║ Office LTSC 2024 ║
echo ║2║ Office LTSC 2021 ║
echo ║3║ Office 2019      ║
echo ╚═╩═════╦════════════╝
choice /C:123 /M "Введіть ╚→ :" /N
set ChoiceOffice=%errorlevel%
if %ChoiceOffice% == 1 goto LTSC2024
if %ChoiceOffice% == 2 goto LTSC2021
if %ChoiceOffice% == 3 goto Office2019


@rem Налаштування для Office LTSC 2024
:LTSC2024
set CID=ef5c8a1f-1356-46fc-984b-634b44e23987
set Channel=PerpetualVL2024
set PID=ProPlus2024Volume
set PIDKEY=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
set OfficeVer=Office LTSC 2024
set ExcludeGroove=0
goto ConfigGen

@rem Налаштування для Office LTSC 2021
:LTSC2021
set CID=c04f0bb9-2868-4356-8632-88c4c1a4870c
set Channel=PerpetualVL2021
set PID=ProPlus2021Volume
set PIDKEY=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
set OfficeVer=Office LTSC 2021
set ExcludeGroove=0
goto ConfigGen

@rem Налаштування для Office 2019
:Office2019
set CID=906df582-99a6-4c42-95e0-a13f220cd505
set Channel=PerpetualVL2019
set PID=ProPlus2019Volume
set PIDKEY=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
set OfficeVer=Office 2019
set ExcludeGroove=1
goto ConfigGen

@rem Генерація XML файлу конфігурації
:ConfigGen
cls
echo Генерація XML файлу конфігурації...
(
    echo ^<Configuration ID="%CID%"^>
    echo   ^<Add OfficeClientEdition="%OCE%" Channel="%Channel%"^>
    echo     ^<Product ID="%PID%" PIDKEY="%PIDKEY%"^>
    echo       ^<Language ID="uk-ua" /^>
    echo       ^<Language ID="en-us" /^>
    echo       ^<ExcludeApp ID="%Access%"/^>
    echo       ^<ExcludeApp ID="%Excel%"/^>
    echo       ^<ExcludeApp ID="Groove"/^>
    echo       ^<ExcludeApp ID="Lync"/^>
    echo       ^<ExcludeApp ID="OneDrive"/^>
    echo       ^<ExcludeApp ID="%OneNote%"/^>
    echo       ^<ExcludeApp ID="%Outlook%"/^>
    echo       ^<ExcludeApp ID="%PowerPoint%"/^>
    echo       ^<ExcludeApp ID="%Publisher%"/^>
    echo       ^<ExcludeApp ID="%Word%"/^>
    echo     ^</Product^>
    echo   ^</Add^>
    echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>
    echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /^>
    echo   ^<Property Name="DeviceBasedLicensing" Value="0" /^>
    echo   ^<Property Name="SCLCacheOverride" Value="0" /^>
    echo   ^<Property Name="AUTOACTIVATE" Value="1" /^>
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

@rem Завантаження Office
:OfficeSetup
setlocal
set URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
set DEST=%~dp0OfficeSetupFiles\
set FILE=OfficeExtracter.exe
@rem Завантаження файлу за допомогою curl
echo Завантаження файлу %FILE%...
curl -L -o %DEST%\%FILE% %URL%
@rem Перевірка, чи файл успішно завантажено
if exist %DEST%\%FILE% (
    echo Файл успішно завантажено в %DEST%%FILE%
) else (
    echo Помилка завантаження файлу.
)
endlocal
timeout 2
cls
start %~dp0OfficeSetupFiles\OfficeExtracter.exe /extract:%~dp0OfficeSetupFiles\ /passive /norestart /quiet
echo Будьласка зачекайте...
timeout 3 && del /f %~dp0OfficeSetupFiles\configuration-Office365-x64.xml && cls
echo Будьласка зачекайте...
timeout 3 && del /f %~dp0OfficeSetupFiles\OfficeExtracter.exe && cls && echo Завантаження %OfficeVer% %OCE%-Біти...
start %~dp0OfficeSetupFiles\setup.exe /configure %~dp0OfficeSetupFiles\Config.xml
timeout 10 && exit

:exit
exit