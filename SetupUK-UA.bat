@chcp 65001
@echo off
@setlocal EnableDelayedExpansion

title ПОТРІБНІ ПРАВА АДМІНІСТРАТОРА
net session >nul
if %errorlevel% neq 0 (cls
echo ^|
echo ^|   ╠══╦═══════════════════════════════╦══╣
echo ^|      ║ ПОТРІБНІ ПРАВА АДМІНІСТРАТОРА ║ 
echo ^|   ╠══╩═══════════════════════════════╩══╣
echo ^| && goto RunScriptAsAdmin
)
goto Start

:RunScriptAsAdmin
cd /d %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
echo ^| Exiting... && timeout 2 >nul && exit

:Start
rmdir /s /q "%TEMP%\OfficeSetupFiles\"
mkdir %TEMP%\OfficeSetupFiles
title Batch Office Installer v1.0.9 від MaximeriX && set DebugMode=0
cls
echo ^|
echo ^|   ╠══╦════════════════════════════════════════════════════════════════════════╦══╣
echo ^|      ║ Batch Office Installer v1.0.9 від                                      ║
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
goto OfficeInfo

:OfficeInfo
cls
echo ^|
echo ^|   ╠══╦═════════════════════════════════════════════════════════════════════════╦══╣
echo ^|      ║ Office включає - Access, Excel, OneDrive, OneNote, Outlook, PowerPoint, ║
echo ^|      ║ Project, Publisher, Teams, Visio and Word. (Можна змінити)              ║
echo ^|      ╠═══╦═══════╗                                                             ║
echo ^|      ║ 1 ║ Добре ║                                                             ║
echo ^|      ║ 2 ║ Вийти ║                                                             ║
echo ^|   ╠══╩═══╩═══════╩══╦══════════════════════════════════════════════════════════╩══╣
echo ^|                     ║
choice /C:123 /M "|   Введіть ваш вибір ╚→ :" /N
set OfficeInfo=%errorlevel%
if %OfficeInfo% == 1 timeout 1 >nul && goto OSArchitectureCheck
if %OfficeInfo% == 2 echo ^| && echo ^| Exiting... && echo ^| && timeout 1 >nul && exit
if %OfficeInfo% == 3 set DebugMode=1 && echo ^| V && timeout 2 >nul && goto OSArchitectureCheck

:OSArchitectureCheck
cls
echo ^|
echo ^|   ╠══╦═══════════════════════════╦══╣
echo ^|      ║ Перевірка типу системи... ║
echo ^|   ╠══╩═══════════════════════════╩══╣
echo ^|
for /f "tokens=2 delims==" %%i in ('wmic os get osarchitecture /value') do (set OSArchitecture=%%i)
if "%OSArchitecture%"=="32-bit" (set OfficeEdition=32
) else (set OfficeEdition=64)
echo ^| ОС %OfficeEdition%-Бітна
timeout 2 >nul && goto OfficeSelect

:OfficeSelect
cls && set Groove=1
echo ^|
echo ^|   ╠══╦════════════════════════════════════════════════╦══╣
echo ^|      ║ Виберіть версію Office, яку хочете встановити. ║
echo ^|      ╠═══╦═══════════════════════════╗                ║
echo ^|      ║ 1 ║ Office LTSC Pro Plus 2024 ║                ║
echo ^|      ║ 2 ║ Office LTSC Standart 2024 ║                ║
echo ^|      ║ 3 ║ Office LTSC Pro Plus 2021 ║                ║
echo ^|      ║ 4 ║ Office LTSC Standart 2021 ║                ║
echo ^|      ║ 7 ║ Office Pro Plus 2016      ║                ║
echo ^|      ║ 8 ║ Office Standart 2016      ║                ║
echo ^|   ╠══╩═══╩══════════╦════════════════╩════════════════╩══╣
echo ^|                     ║
choice /C:12345678 /M "|   Введіть ваш вибір ╚→ :" /N
set SelectVer=%errorlevel%
if %SelectVer% == 1 (set UpdateChannel=PerpetualVL2024
    set ProductID=ProPlus2024Volume
    set ProductKey=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
    set OfficeVersion=Office LTSC Pro Plus 2024
    set ProductIDVS=VisioPro2024Volume
    set ProductKeyVS=B7TN8-FJ8V3-7QYCP-HQPMV-YY89G
    set ProductIDPR=ProjectPro2024Volume
    set ProductKeyPR=FQQ23-N4YCY-73HQ3-FM9WC-76HF4
) else if %SelectVer% == 2 (set UpdateChannel=PerpetualVL2024
    set ProductID=Standard2024Volume
    set ProductKey=V28N4-JG22K-W66P8-VTMGK-H6HGR
    set OfficeVersion=Office LTSC Standart 2024
    set ProductIDVS=VisioStd2024Volume
    set ProductKeyVS=JMMVY-XFNQC-KK4HK-9H7R3-WQQTV
    set ProductIDPR=ProjectStd2024Volume
    set ProductKeyPR=PD3TT-NTHQQ-VC7CY-MFXK3-G87F8
) else if %SelectVer% == 3 (set UpdateChannel=PerpetualVL2021
    set ProductID=ProPlus2021Volume
    set ProductKey=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
    set OfficeVersion=Office LTSC Pro Plus 2021
    set ProductIDVS=VisioPro2021Volume
    set ProductKeyVS=KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
    set ProductIDPR=ProjectPro2021Volume
    set ProductKeyPR=FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
) else if %SelectVer% == 4 (set UpdateChannel=PerpetualVL2021
    set ProductID=Standard2021Volume
    set ProductKey=KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3
    set OfficeVersion=Office LTSC Standart 2021
    set ProductIDVS=VisioStd2021Volume
    set ProductKeyVS=MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
    set ProductIDPR=ProjectStd2021Volume
    set ProductKeyPR=J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
) else if %SelectVer% == 5 (set UpdateChannel=PerpetualVL2019
    set ProductID=ProPlus2019Volume
    set ProductKey=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
    set OfficeVersion=Office Pro Plus 2019
    set ProductIDVS=VisioPro2019Volume
    set ProductKeyVS=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
    set ProductIDPR=ProjectPro2019Volume
    set ProductKeyPR=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
) else if %SelectVer% == 6 (set UpdateChannel=PerpetualVL2019
    set ProductID=Standard2019Volume
    set ProductKey=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
    set OfficeVersion=Office Standart 2019
    set ProductIDVS=VisioStd2019Volume
    set ProductKeyVS=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
    set ProductIDPR=ProjectStd2019Volume
    set ProductKeyPR=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
) else if %SelectVer% == 7 (set UpdateChannel=Broad
    set ProductID=ProPlusRetail
    set ProductKey=CYC3N-BHX8G-QJVJV-H2WWP-BTDRB
    set OfficeVersion=Office Pro Plus 2016
    set ProductIDVS=VisioProXVolume
    set ProductKeyVS=69WXN-MBYV6-22PQG-3WGHK-RM6XC
    set ProductIDPR=ProjectProXVolume
    set ProductKeyPR=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9
) else if %SelectVer% == 8 (set UpdateChannel=Broad
    set ProductID=StandardRetail
    set ProductKey=PCCXN-7MKB3-F986V-V6HV4-CR4MR
    set OfficeVersion=Office Standart 2016
    set ProductIDVS=VisioStdXVolume
    set ProductKeyVS=NY48V-PPYYH-3F4PX-XJRKJ-W4423
    set ProductIDPR=ProjectStdXVolume
    set ProductKeyPR=D8NRQ-JTYM3-7J2DX-646CT-6836M
)
timeout 2 >nul && goto AppsInstall 

:AppsInstall
cls && set Access=0 && set Excel=0 && set OneDrive=0 && set OneNote=0 && set Outlook=0 && set PowerPoint=0 && set Project=0 && set Publisher=0 && set Teams=0 && set Visio=0 && set Word=0
echo ^|
echo ^|   ╠══╦══════════════════════════════════════════════╦══╣
echo ^|      ║ Виберіть програми, які ви хочете встановити. ║
echo ^|      ╠═══╦════════════╦════╦══════════════╗         ║
echo ^|      ║ 1 ║ Access     ║ 7  ║ Project      ║         ║
echo ^|      ║ 2 ║ Excel      ║ 8  ║ Publisher    ║         ║
echo ^|      ║ 3 ║ OneDrive   ║ 9  ║ Teams        ║         ║
echo ^|      ║ 4 ║ OneNote    ║ 10 ║ Visio        ║         ║
echo ^|      ║ 5 ║ Outlook    ║ 11 ║ Word         ║         ║
echo ^|      ║ 6 ║ PowerPoint ║ A  ║ Залишити всі ║         ║
echo ^|   ╠══╩═══╩════════════╩════╩═══╦══════════╩═════════╩══╣
echo ^|                                ║  
set /p AppsInstall="|   Введіть ваш вибір (1 4 тощо) ╚→ : "
for %%i in (%AppsInstall%) do (
    if %%i==1 (set Access=1
    ) else if %%i==2 (set Excel=1
    ) else if %%i==3 (set OneDrive=1
    ) else if %%i==4 (set OneNote=1
    ) else if %%i==5 (set Outlook=1
    ) else if %%i==6 (set PowerPoint=1
    ) else if %%i==7 (set Project=1
    ) else if %%i==8 (set Publisher=1
    ) else if %%i==9 (set Teams=1
    ) else if %%i==10 (set Visio=1
    ) else if %%i==11 (set Word=1
    ) else if %%i==A (set Access=1 && set Excel=1 && set OneDrive=1 && set OneNote=1 && set Outlook=1 && set PowerPoint=1 && set Project=1 && set Publisher=1 && set Teams=1 && set Visio=1 && set Word=1
    ) else if %%i==a (set Access=1 && set Excel=1 && set OneDrive=1 && set OneNote=1 && set Outlook=1 && set PowerPoint=1 && set Project=1 && set Publisher=1 && set Teams=1 && set Visio=1 && set Word=1
    ) else if %%i==А (set Access=1 && set Excel=1 && set OneDrive=1 && set OneNote=1 && set Outlook=1 && set PowerPoint=1 && set Project=1 && set Publisher=1 && set Teams=1 && set Visio=1 && set Word=1
    ) else if %%i==а (set Access=1 && set Excel=1 && set OneDrive=1 && set OneNote=1 && set Outlook=1 && set PowerPoint=1 && set Project=1 && set Publisher=1 && set Teams=1 && set Visio=1 && set Word=1
    ) else (
        echo ^| 
        echo ^| Неправильний вибір: %%i
        echo ^| Вихід... 
        echo ^| && timeout 3 >nul && exit
    )
)
if %DebugMode% == 1 (timeout 1 >nul && goto Debug
) else (timeout 1 >nul && goto ConfigGen)

:Debug
cls && echo ^| Current variable values: && echo ^| Access: %Access% && echo ^| Excel: %Excel% && echo ^| Groove: %Groove% && echo ^| OneDrive: %OneDrive% && echo ^| OneNote: %OneNote% && echo ^| Outlook: %Outlook% && echo ^| PowerPoint: %PowerPoint% && echo ^| Project: %Project% && echo ^| Publisher: %Publisher% && echo ^| Teams: %Teams% && echo ^| Visio: %Visio% && echo ^| Word: %Word% && echo ^| && echo ^| ProductIDPR: %ProductIDPR% && echo ^| ProductKeyPR: %ProductKeyPR% && echo ^| ProductIDVS: %ProductIDVS% && echo ^| ProductKeyVS: %ProductKeyVS% && echo ^| ConfigurationID: %ConfigurationID% && echo ^| UpdateChannel: %UpdateChannel% && echo ^| ProductID: %ProductID% && echo ^| ProductKey: %ProductKey% && echo ^| OfficeVersion: %OfficeVersion% && echo ^| Bits: %OfficeEdition% && echo ^| && echo ^| Press any key to continue...
pause >nul && goto ConfigGen

:ConfigGen
cls
set ConfigPath=%TEMP%\OfficeSetupFiles\Config.xml
echo ^| 
echo ^|   ╠══╦═════════════════════════════════════╦══╣
echo ^|      ║ Генерація XML файлу конфігурації... ║
echo ^|   ╠══╩═════════════════════════════════════╩══╣
echo ^|
(
    echo ^<Configuration^>
    echo   ^<Add OfficeClientEdition="%OfficeEdition%" Channel="%UpdateChannel%"^>
    echo     ^<Product ID="%ProductID%" PIDKEY="%ProductKey%"^>
    echo       ^<Language ID="MatchOS" /^>
    echo       ^<Language ID="uk-ua" /^>
    echo       ^<Language ID="en-us" /^>
    if %Access%==0 echo       ^<ExcludeApp ID="Access"/^>
    if %Excel%==0 echo       ^<ExcludeApp ID="Excel"/^>
    if %Groove%==0 echo       ^<ExcludeApp ID="Groove"/^>
    echo       ^<ExcludeApp ID="Lync"/^>
    if %OneDrive%==0 echo       ^<ExcludeApp ID="OneDrive"/^>
    if %OneNote%==0 echo       ^<ExcludeApp ID="OneNote"/^>
    if %Outlook%==0 echo       ^<ExcludeApp ID="Outlook"/^>
    if %PowerPoint%==0 echo       ^<ExcludeApp ID="PowerPoint"/^>
    if %Publisher%==0 echo       ^<ExcludeApp ID="Publisher"/^>
    if %Word%==0 echo       ^<ExcludeApp ID="Word"/^>
    echo     ^</Product^>
    if %Project%==1 (echo     ^<Product ID="%ProductIDPR%" PIDKEY="%ProductKeyPR%"^>
        echo       ^<Language ID="MatchOS" /^>
        echo       ^<Language ID="uk-ua" /^>
        echo       ^<Language ID="en-us" /^>
        if %Access%==0 echo       ^<ExcludeApp ID="Access"/^>
        if %Excel%==0 echo       ^<ExcludeApp ID="Excel"/^>
        if %Groove%==0 echo       ^<ExcludeApp ID="Groove"/^>
        echo       ^<ExcludeApp ID="Lync"/^>
        if %OneDrive%==0 echo       ^<ExcludeApp ID="OneDrive"/^>
        if %OneNote%==0 echo       ^<ExcludeApp ID="OneNote"/^>
        if %Outlook%==0 echo       ^<ExcludeApp ID="Outlook"/^>
        if %PowerPoint%==0 echo       ^<ExcludeApp ID="PowerPoint"/^>
        if %Publisher%==0 echo       ^<ExcludeApp ID="Publisher"/^>
        if %Word%==0 echo       ^<ExcludeApp ID="Word"/^>
        echo     ^</Product^>)
    if %Visio%==1 (echo     ^<Product ID="%ProductIDVS%" PIDKEY="%ProductKeyVS%"^>
        echo       ^<Language ID="MatchOS" /^>
        echo       ^<Language ID="uk-ua" /^>
        echo       ^<Language ID="en-us" /^>
        if %Access%==0 echo       ^<ExcludeApp ID="Access"/^>
        if %Excel%==0 echo       ^<ExcludeApp ID="Excel"/^>
        if %Groove%==0 echo       ^<ExcludeApp ID="Groove"/^>
        echo       ^<ExcludeApp ID="Lync"/^>
        if %OneDrive%==0 echo       ^<ExcludeApp ID="OneDrive"/^>
        if %OneNote%==0 echo       ^<ExcludeApp ID="OneNote"/^>
        if %Outlook%==0 echo       ^<ExcludeApp ID="Outlook"/^>
        if %PowerPoint%==0 echo       ^<ExcludeApp ID="PowerPoint"/^>
        if %Publisher%==0 echo       ^<ExcludeApp ID="Publisher"/^>
        if %Word%==0 echo       ^<ExcludeApp ID="Word"/^>
        echo     ^</Product^>)
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
) > %ConfigPath%
if exist %ConfigPath% (echo ^| Конфігурація збережена до ^> %ConfigPath%
) else (
    echo ^| Помилка збереження конфігурації
    echo ^| Вихід...
    timeout 7 >nul && exit
)
if %DebugMode% == 1 (pause >nul && goto FilesDownload
) else (timeout 1 >nul && goto FilesDownload)

:FilesDownload
set PATH=%TEMP%\OfficeSetupFiles\
set ExtractorPath=%TEMP%\OfficeSetupFiles\OfficeExtracter.exe
set TeamsPath=%TEMP%\OfficeSetupFiles\MSTeamsSetup.exe
set SetupPath=%TEMP%\OfficeSetupFiles\setup.exe
if %DebugMode% == 0 (if %Teams%==1 (cls
        echo ^|
        echo ^|   ╠══╦══════════════════════════════════╦══╣
        echo ^|      ║ Завантаження MSTeamsSetup.exe... ║
        echo ^|   ╠══╩══════════════════════════════════╩══╣
        echo ^|
        curl -L -s -o %TeamsPath% https://go.microsoft.com/fwlink/?linkid=2281613&clcid=0x409
        if exist %TeamsPath% (echo ^| Успішно завантажено до %TeamsPath%
            timeout 1 >nul
            start %TeamsPath%
            cls
            echo ^|
            echo ^|   ╠══╦═══════════════════════╦══╣
            echo ^|      ║ Завантаження Teams... ║
            echo ^|   ╠══╩═══════════════════════╩══╣
            echo ^| && timeout 1 >nul && goto TeamsCheckLoop
        ) else (echo ^| Помилка завантаження MSTeamsSetup.exe.
            echo ^| Вихід...
            timeout 7 >nul && exit)
    ) else (timeout 1 >nul && goto OfficeExtracterDownload)
) else (timeout 1 >nul && goto OfficeExtracterDownload)

:TeamsCheckLoop
tasklist /fi "imagename eq MSTeamsSetup.exe" | find /i "MSTeamsSetup.exe" >nul
if errorlevel 1 (taskkill /f /im ms-teams.exe >nul
    timeout 2 >nul && del /f %TeamsPath%
    goto OfficeExtracterDownload
) else (timeout 1 >nul && goto TeamsCheckLoop)

:OfficeExtracterDownload
cls
echo ^|
echo ^|   ╠══╦═════════════════════════════════════╦══╣
echo ^|      ║ Завантаження OfficeExtracter.exe... ║
echo ^|   ╠══╩═════════════════════════════════════╩══╣
echo ^|
curl -L -s -o %ExtractorPath% https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
if exist %ExtractorPath% (echo ^| Успішно завантажено до %ExtractorPath%
    timeout 1 >nul && goto Extracting
) else (echo ^| Помилка завантаження OfficeExtracter.exe
    echo ^| Вихід...
    timeout 10 >nul && exit)

:Extracting
cls
echo ^|
echo ^|   ╠══╦═════════════════════════════╦══╣
echo ^|      ║ Екстракція файлів Office... ║
echo ^|   ╠══╩═════════════════════════════╩══╣
echo ^|
start %ExtractorPath% /extract:%PATH% /passive /norestart /quiet && goto ExtractingLoop

:ExtractingLoop
tasklist /fi "imagename eq OfficeExtracter.exe" | find /i "OfficeExtracter.exe" >nul
if errorlevel 1 (
    timeout 1 >nul && del /f %TEMP%\OfficeSetupFiles\configuration-Office365-x64.xml
    timeout 1 >nul && del /f %ExtractorPath%
    goto OfficeInstallerStart
) else (
    timeout 1 >nul && goto ExtractingLoop)

:OfficeInstallerStart
cls
echo ^|
echo ^|   ╠══╦══════════════════════════════╦══╣
echo ^|      ║ Запуск інсталятора Office... ║
echo ^|   ╠══╩══════════════════════════════╩══╣
echo ^|
echo ^| Інсталятор для %OfficeVersion% %OfficeEdition%-Біт запущено...
if %DebugMode% == 1 (echo ^| ** && echo start %SetupPath% /configure %ConfigPath% && echo ^| **
) else (start %SetupPath% /configure %ConfigPath%)
timeout 3 >nul
echo ^|
echo ^| Дякую за використання мого скрипту. Будь ласка, підтримайте мене на Ko-fi: https://ko-fi.com/MaximeriX
echo ^| Натисніть 1, щоб відкрити посилання
echo ^| Натисніть 2, щоб вийти
choice /C:12 /M "| >" /N
set Donation=%errorlevel%
if %Donation% == 1 (start https://ko-fi.com/MaximeriX
) else ( echo ^| Вихід... && timeout 2 >nul && exit)
@endlocal