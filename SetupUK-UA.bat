@chcp 65001
@echo off
@setlocal EnableDelayedExpansion
title ПОТРІБНІ ПРАВА АДМІНІСТРАТОРА

@rem Перевірка прав адміністратора
net session
cls
echo ^|
echo ^|   ╠══╦═══════════════════════════════╦══╣
echo ^|      ║ ПОТРІБНІ ПРАВА АДМІНІСТРАТОРА ║ 
echo ^|   ╠══╩═══════════════════════════════╩══╣
echo ^|
if %errorlevel% neq 0 goto runadmin
goto adminstart

@rem Запуск командного рядка від імені адміністратора
:runadmin
CD /d %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
echo ^| Вихід... && timeout 2 >nul && exit

@rem Початок скрипту
:adminstart
:start
rmdir /s /q "%TEMP%\OfficeSetupFiles\"
mkdir %TEMP%\OfficeSetupFiles
title Simple Office Installer від MaximeriX
set Debug=0
cls
echo ^|
echo ^|   ╠══╦════════════════════════════════════════════════════════════════════════╦══╣
echo ^|      ║ Simple Office Installer від                                            ║
echo ^|      ╠════════════════════════════════════════════════════════════════════════╣
echo ^|      ║                                                                        ║
echo ^|      ║  ███╗   ███╗ █████╗ ██╗  ██╗██╗███╗   ███╗███████╗██████╗ ██╗██╗  ██╗  ║
echo ^|      ║  ████╗ ████║██╔══██╗╚██╗██╔╝██║████╗ ████║██╔════╝██╔══██╗██║╚██╗██╔╝  ║
echo ^|      ║  ██╔████╔██║███████║ ╚███╔╝ ██║██╔████╔██║█████╗  ██████╔╝██║ ╚███╔╝   ║
echo ^|      ║  ██║╚██╔╝██║██╔══██║ ██╔██╗ ██║██║╚██╔╝██║██╔══╝  ██╔══██╗██║ ██╔██╗   ║
echo ^|      ║  ██║ ╚═╝ ██║██║  ██║██╔╝╚██╗██║██║ ╚═╝ ██║███████╗██║  ██║██║██╔╝╚██╗  ║
echo ^|      ║  ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝╚═╝╚═╝  ╚═╝  ║
echo ^|      ║                                                                        ║
echo ^|   ╠══╩════════════════════════════════════════════════════════════════════════╩══╣
echo ^|      
timeout 2 >nul

goto Office
:Office
cls
echo ^|
echo ^|   ╠══╦════════════════════════════════════════════════════════════════════════════════════════════════╦══╣
echo ^|      ║ Office - включає Access, Excel, OneNote, Outlook, PowerPoint, Project, Publisher, Visio, Word. ║
echo ^|      ╠═══╦═══════╗                                                                                    ║
echo ^|      ║ 1 ║ Добре ║                                                                                    ║
echo ^|      ║ 2 ║ Вийти ║                                                                                    ║
echo ^|   ╠══╩═══╩═══════╩══╦═════════════════════════════════════════════════════════════════════════════════╩══╣
echo ^|                     ║
choice /C:123 /M "|   Введіть ваш вибір ╚→ :" /N
set UserChoice=%errorlevel%
if %UserChoice% == 1 timeout 1 >nul && goto ExcludeApps
if %UserChoice% == 2 echo ^| && echo ^| Вихід... && echo ^| && timeout 1 >nul && exit
if %UserChoice% == 3 set Debug=1 && echo ^| V && timeout 2 >nul && goto ExcludeApps

@rem Вибір програм для виключення
:ExcludeApps
set Access=1
set Excel=1
set OneNote=1
set Outlook=1
set PowerPoint=1
set Project=1
set Publisher=1
set Visio=1
set Word=1
cls
echo ^|
echo ^|   ╠══╦════════════════════════════════════════════════════╦══╣
echo ^|      ║ Виберіть програми, які ви не хочете завантажувати. ║
echo ^|      ╠═══╦══════════════╗                                 ║
echo ^|      ║ 1 ║ Access       ║                                 ║
echo ^|      ║ 2 ║ Excel        ║                                 ║
echo ^|      ║ 3 ║ OneNote      ║                                 ║
echo ^|      ║ 4 ║ Outlook      ║                                 ║
echo ^|      ║ 5 ║ PowerPoint   ║                                 ║
echo ^|      ║ 6 ║ Project      ║                                 ║
echo ^|      ║ 7 ║ Publisher    ║                                 ║
echo ^|      ║ 8 ║ Visio        ║                                 ║
echo ^|      ║ 9 ║ Word         ║                                 ║
echo ^|      ║ A ║ Залишити всі ║                                 ║
echo ^|   ╠══╩═══╩══════════════╩══════╦══════════════════════════╩══╣
echo ^|                                ║   
set "excludeApps="
set /p input="|   Введіть ваш вибір (1 4 тощо) ╚→ : "

for %%i in (%input%) do (
    if %%i==1 (
        set Access=0
    ) else if %%i==2 (
        set Excel=0
    ) else if %%i==3 (
        set OneNote=0
    ) else if %%i==4 (
        set Outlook=0
    ) else if %%i==5 (
        set PowerPoint=0
    ) else if %%i==6 (
        set Project=0
    ) else if %%i==7 (
        set Publisher=0
    ) else if %%i==8 (
        set Visio=0
    ) else if %%i==9 (
        goto Word=0
    ) else if %%i==A (
        goto bitcheck
    ) else if %%i==a (
        goto bitcheck
    ) else if %%i==А (
        goto bitcheck
    ) else if %%i==а (
        goto bitcheck
    ) else (
        echo ^| 
        echo ^| Неправильний вибір: %%i
        echo ^| Вихід... 
        echo ^| && timeout 3 >nul && exit
    )
)
goto bitcheck

@rem Перевірка архітектури системи
:bitcheck
cls
echo ^|
echo ^|   ╠══╦═══════════════════════════╦══╣
echo ^|      ║ Перевірка типу системи... ║
echo ^|   ╠══╩═══════════════════════════╩══╣
echo ^|
for /f "tokens=2 delims==" %%i in ('wmic os get osarchitecture /value') do (
 set architecture=%%i
)
if "%architecture%"=="32-bit" (
    set OfficeEdition=32
) else (
    set OfficeEdition=64
)
echo ^| ОС %OfficeEdition%-Бітна
timeout 2 >nul && goto OfficeSelect

@rem Вибір версії Office
:OfficeSelect
cls
set Groove=1
echo ^|
echo ^|   ╠══╦══════════════════════════════════════════════════╦══╣
echo ^|      ║ Виберіть версію Office, яку ви хочете встановити ║
echo ^|      ╠═══╦══════════════════╗                           ║
echo ^|      ║ 1 ║ Office LTSC 2024 ║                           ║
echo ^|      ║ 2 ║ Office LTSC 2021 ║                           ║
echo ^|      ║ 3 ║ Office 2019      ║                           ║
echo ^|   ╠══╩═══╩══════════╦═══════╩═══════════════════════════╩══╣
echo ^|                     ║
choice /C:123 /M "|   Введіть ваш вибір ╚→ :" /N
set OfficeChoice=%errorlevel%
if %OfficeChoice% == 1 timeout 1 >nul && goto LTSC2024
if %OfficeChoice% == 2 timeout 1 >nul && goto LTSC2021
if %OfficeChoice% == 3 timeout 1 >nul && set Groove=0 && goto Office2019

@rem Налаштування для Office LTSC 2024
:LTSC2024
set ProductIDPR=ProjectPro2024Volume
set ProductKeyPR=FQQ23-N4YCY-73HQ3-FM9WC-76HF4
set ProductIDVS=VisioPro2024Volume
set ProductKeyVS=B7TN8-FJ8V3-7QYCP-HQPMV-YY89G
set ConfigurationID=ef5c8a1f-1356-46fc-984b-634b44e23987
set UpdateChannel=PerpetualVL2024
set ProductID=ProPlus2024Volume
set ProductKey=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
set OfficeVersion=Office LTSC 2024
if %Debug% == 1 (
    goto Debug
) else (
    goto ConfigGen
)

@rem Налаштування для Office LTSC 2021
:LTSC2021
set ProductIDPR=ProjectPro2021Volume
set ProductKeyPR=FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
set ProductIDVS=VisioPro2021Volume
set ProductKeyVS=KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
set ConfigurationID=c04f0bb9-2868-4356-8632-88c4c1a4870c
set UpdateChannel=PerpetualVL2021
set ProductID=ProPlus2021Volume
set ProductKey=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
set OfficeVersion=Office LTSC 2021
if %Debug% == 1 (
    goto Debug
) else (
    goto ConfigGen
)

@rem Налаштування для Office 2019
:Office2019
set ProductIDPR=ProjectPro2019Volume
set ProductKeyPR=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
set ProductIDVS=VisioPro2019Volume
set ProductKeyVS=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
set ConfigurationID=906df582-99a6-4c42-95e0-a13f220cd505
set UpdateChannel=PerpetualVL2019
set ProductID=ProPlus2019Volume
set ProductKey=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
set OfficeVersion=Office 2019
if %Debug% == 1 (
    goto Debug
) else (
    goto ConfigGen
)

:Debug
cls
echo ^| Поточні значення змінних:
echo ^| Access: %Access%
echo ^| Excel: %Excel%
echo ^| Groove: %Groove%
echo ^| OneNote: %OneNote%
echo ^| Outlook: %Outlook%
echo ^| PowerPoint: %PowerPoint%
echo ^| Publisher: %Publisher%
echo ^| Word: %Word%
echo ^| Project: %Project%
echo ^| Visio: %Visio%
echo ^|
echo ^| ProductIDPR: %ProductIDPR%
echo ^| ProductKeyPR: %ProductKeyPR%
echo ^| ProductIDVS: %ProductIDVS%
echo ^| ProductKeyVS: %ProductKeyVS%
echo ^| ConfigurationID: %ConfigurationID%
echo ^| UpdateChannel: %UpdateChannel%
echo ^| ProductID: %ProductID%
echo ^| ProductKey: %ProductKey%
echo ^| OfficeVersion: %OfficeVersion%
echo ^|
pause >nul
goto ConfigGen

@rem Генерація XML файлу конфігурації
:ConfigGen
cls
echo ^| 
echo ^|   ╠══╦═════════════════════════════════════╦══╣
echo ^|      ║ Генерація XML файлу конфігурації... ║
echo ^|   ╠══╩═════════════════════════════════════╩══╣
echo ^|
(
    echo ^<Configuration ID="%ConfigurationID%"^>
    echo   ^<Add OfficeClientEdition="%OfficeEdition%" Channel="%UpdateChannel%"^>
    echo     ^<Product ID="%ProductID%" PIDKEY="%ProductKey%"^>
    echo       ^<Language ID="uk-ua" /^>
    echo       ^<Language ID="en-us" /^>
    if %Access%==0 echo       ^<ExcludeApp ID="Access"/^>
    if %Excel%==0 echo       ^<ExcludeApp ID="Excel"/^>
    if %Groove%==0 echo       ^<ExcludeApp ID="Groove"/^>
    echo       ^<ExcludeApp ID="Lync"/^>
    echo       ^<ExcludeApp ID="OneDrive"/^>
    if %OneNote%==0 echo       ^<ExcludeApp ID="OneNote"/^>
    if %Outlook%==0 echo       ^<ExcludeApp ID="Outlook"/^>
    if %PowerPoint%==0 echo       ^<ExcludeApp ID="PowerPoint"/^>
    if %Publisher%==0 echo       ^<ExcludeApp ID="Publisher"/^>
    if %Word%==0 echo       ^<ExcludeApp ID="Word"/^>
    echo     ^</Product^>
    if %Project%==1 (
        echo     ^<Product ID="%ProductIDPR%" PIDKEY="%ProductKeyPR%"^>
        echo       ^<Language ID="uk-ua" /^>
        echo       ^<Language ID="en-us" /^>
        if %Access%==0 echo       ^<ExcludeApp ID="Access"/^>
        if %Excel%==0 echo       ^<ExcludeApp ID="Excel"/^>
        if %Groove%==0 echo       ^<ExcludeApp ID="Groove"/^>
        echo       ^<ExcludeApp ID="Lync"/^>
        echo       ^<ExcludeApp ID="OneDrive"/^>
        if %OneNote%==0 echo       ^<ExcludeApp ID="OneNote"/^>
        if %Outlook%==0 echo       ^<ExcludeApp ID="Outlook"/^>
        if %PowerPoint%==0 echo       ^<ExcludeApp ID="PowerPoint"/^>
        if %Publisher%==0 echo       ^<ExcludeApp ID="Publisher"/^>
        if %Word%==0 echo       ^<ExcludeApp ID="Word"/^>
        echo     ^</Product^>
    )
    if %Visio%==1 (
        echo     ^<Product ID="%ProductIDVS%" PIDKEY="%ProductKeyVS%"^>
        echo       ^<Language ID="uk-ua" /^>
        echo       ^<Language ID="en-us" /^>
        if %Access%==0 echo       ^<ExcludeApp ID="Access"/^>
        if %Excel%==0 echo       ^<ExcludeApp ID="Excel"/^>
        if %Groove%==0 echo       ^<ExcludeApp ID="Groove"/^>
        echo       ^<ExcludeApp ID="Lync"/^>
        echo       ^<ExcludeApp ID="OneDrive"/^>
        if %OneNote%==0 echo       ^<ExcludeApp ID="OneNote"/^>
        if %Outlook%==0 echo       ^<ExcludeApp ID="Outlook"/^>
        if %PowerPoint%==0 echo       ^<ExcludeApp ID="PowerPoint"/^>
        if %Publisher%==0 echo       ^<ExcludeApp ID="Publisher"/^>
        if %Word%==0 echo       ^<ExcludeApp ID="Word"/^>
        echo     ^</Product^>
    )
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
) > %TEMP%\OfficeSetupFiles\Config.xml

echo ^| Конфігурація збережена в ^> %TEMP%\OfficeSetupFiles\Config.xml
timeout 1 >nul
goto OfficeExtracterDownload

@rem Завантаження Office
:OfficeExtracterDownload
set PATH=%TEMP%\OfficeSetupFiles\
set ExtractorPath=%TEMP%\OfficeSetupFiles\OfficeExtracter.exe
set SetupPath=%TEMP%\OfficeSetupFiles\setup.exe
setlocal
@rem Завантаження файлу за допомогою curl
cls
echo ^|
echo ^|   ╠══╦═════════════════════════════════════╦══╣
echo ^|      ║ Завантаження OfficeExtracter.exe... ║
echo ^|   ╠══╩═════════════════════════════════════╩══╣
echo ^|
curl -L -s -o %ExtractorPath% https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Перевірка, чи файл був успішно завантажений
if exist %ExtractorPath% (
    echo ^| Успішно завантажено до %PATH%
    timeout 1 >nul
) else (
    echo Помилка завантаження OfficeExtracter.exe
    timeout 10 >nul && exit
)
endlocal
goto Extracting

:Extracting
cls
echo ^|
echo ^|   ╠══╦═════════════════════════════╦══╣
echo ^|      ║ Екстракція файлів Office... ║
echo ^|   ╠══╩═════════════════════════════╩══╣
echo ^|
start %ExtractorPath% /extract:%PATH% /passive /norestart /quiet
timeout 2 >nul && del /f %TEMP%\OfficeSetupFiles\OfficeExtracter.exe
timeout 2 >nul && del /f %TEMP%\OfficeSetupFiles\configuration-Office365-x64.xml
goto OfficeInstallerStart

:OfficeInstallerStart
cls
echo ^|
echo ^|   ╠══╦══════════════════════════════╦══╣
echo ^|      ║ Запуск інсталятора Office... ║
echo ^|   ╠══╩══════════════════════════════╩══╣
echo ^|
echo ^| Інсталятор для %OfficeVersion% %OfficeEdition%-Біт запущено...
start %SetupPath% /configure %PATH%Config.xml
timeout 2 >nul
echo ^|
echo ^| Дякуємо за використання мого скрипту. Будь ласка, підтримайте мене на Ko-fi: https://ko-fi.com/MaximeriX
echo ^| Натисніть 1, щоб відкрити посилання
echo ^| Натисніть 2, щоб вийти
choice /C:12 /M "| >" /N
set Donation=%errorlevel%
if %Donation% == 1 start https://ko-fi.com/MaximeriX (
) else ( 
echo ^| Вихід... && timeout 2 >nul && exit
)
@endlocal