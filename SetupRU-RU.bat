@chcp 65001
@echo off
@setlocal EnableDelayedExpansion
title ТРЕБУЮТСЯ ПРАВА АДМИНИСТРАТОРА

@rem Проверка прав администратора
net session
cls
echo ^|
echo ^|   ╠══╦════════════════════════════════╦══╣
echo ^|      ║ ТРЕБУЮТСЯ ПРАВА АДМИНИСТРАТОРА ║ 
echo ^|   ╠══╩════════════════════════════════╩══╣
echo ^|
if %errorlevel% neq 0 goto runadmin
goto adminstart

@rem Запуск командной строки от имени администратора
:runadmin
CD /d %~dp0
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
echo ^| Выход... && timeout 2 >nul && exit

@rem Начало скрипта
:adminstart
:start
rmdir /s /q "%TEMP%\OfficeSetupFiles\"
mkdir %TEMP%\OfficeSetupFiles
title Simple Office Installer от MaximeriX
cls
echo ^|
echo ^|   ╠══╦════════════════════════════════════════════════════════════════════════╦══╣
echo ^|      ║ Simple Office Installer от                                             ║
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
echo ^|   ╠══╦══════════════════════════════════════════════════════════════════════════════════════════════════╦══╣
echo ^|      ║ Office - включает Access, Excel, OneNote, Outlook, PowerPoint, Publisher, Word. (Можно изменить) ║
echo ^|      ╠═══╦═══════╗                                                                                      ║
echo ^|      ║ 1 ║ Окей  ║                                                                                      ║
echo ^|      ║ 2 ║ Выход ║                                                                                      ║
echo ^|   ╠══╩═══╩═══════╩══╦═══════════════════════════════════════════════════════════════════════════════════╩══╣
echo ^|                     ║
choice /C:12 /M "|   Введите ваш выбор ╚→ :" /N
set UserChoice=%errorlevel%
if %UserChoice% == 1 timeout 1 >nul && goto ExcludeApps
if %UserChoice% == 2 echo ^| && echo ^| Выход... && echo ^| && timeout 1 >nul && exit

@rem Выбор приложений для исключения
:ExcludeApps
cls
echo ^|
echo ^|   ╠══╦══════════════════════════════════════════════════════╦══╣
echo ^|      ║ Выберите приложения, которые вы не хотите скачивать. ║
echo ^|      ╠═══╦══════════════╗                                   ║
echo ^|      ║ 1 ║ Access       ║                                   ║
echo ^|      ║ 2 ║ Excel        ║                                   ║
echo ^|      ║ 3 ║ OneNote      ║                                   ║
echo ^|      ║ 4 ║ Outlook      ║                                   ║
echo ^|      ║ 5 ║ PowerPoint   ║                                   ║
echo ^|      ║ 6 ║ Publisher    ║                                   ║
echo ^|      ║ 7 ║ Word         ║                                   ║
echo ^|      ║ 8 ║ Оставить всё ║                                   ║
echo ^|   ╠══╩═══╩══════════════╩════════╦══════════════════════════╩══╣
echo ^|                                  ║   
set "excludeApps="
set /p input="|   Введите ваши выборы (1 4 т.д.) ╚→ : "

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
    ) else if %%i==8 (
        goto bitcheck
    ) else (
        echo ^| 
        echo ^| Неверный выбор: %%i
        echo ^| Выход... 
        echo ^| && timeout 3 >nul && exit
    )
)
goto bitcheck

@rem Проверка архитектуры системы
:bitcheck
cls
echo ^|
echo ^|   ╠══╦══════════════════════════╦══╣
echo ^|      ║ Проверка типа системы... ║
echo ^|   ╠══╩══════════════════════════╩══╣
echo ^|
for /f "tokens=2 delims==" %%i in ('wmic os get osarchitecture /value') do (
 set architecture=%%i
)
if "%architecture%"=="32-bit" (
    set OfficeEdition=32
) else (
    set OfficeEdition=64
)
echo ^| ОС %OfficeEdition%-Битная
timeout 2 >nul && goto OfficeSelect

@rem Выбор версии Office
:OfficeSelect
cls
echo ^|
echo ^|   ╠══╦══════════════════════════════════════════════════════╦══╣
echo ^|      ║ Выберите версию Office, которую вы хотите установить ║
echo ^|      ╠═══╦══════════════════╗                               ║
echo ^|      ║ 1 ║ Office LTSC 2024 ║                               ║
echo ^|      ║ 2 ║ Office LTSC 2021 ║                               ║
echo ^|      ║ 3 ║ Office 2019      ║                               ║
echo ^|   ╠══╩═══╩══════════╦═══════╩═══════════════════════════════╩══╣
echo ^|                     ║
choice /C:123 /M "|   Введите ваш выбор ╚→ :" /N
set OfficeChoice=%errorlevel%
if %OfficeChoice% == 1 timeout 1 >nul && goto LTSC2024
if %OfficeChoice% == 2 timeout 1 >nul && goto LTSC2021
if %OfficeChoice% == 3 timeout 1 >nul && goto Office2019

@rem Настройки для Office LTSC 2024
:LTSC2024
set ConfigurationID=ef5c8a1f-1356-46fc-984b-634b44e23987
set UpdateChannel=PerpetualVL2024
set ProductID=ProPlus2024Volume
set ProductKey=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB
set OfficeVersion=Office LTSC 2024
goto ConfigGen

@rem Наст ```batch
ройки для Office LTSC 2021
:LTSC2021
set ConfigurationID=c04f0bb9-2868-4356-8632-88c4c1a4870c
set UpdateChannel=PerpetualVL2021
set ProductID=ProPlus2021Volume
set ProductKey=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
set OfficeVersion=Office LTSC 2021
goto ConfigGen

@rem Настройки для Office 2019
:Office2019
set ConfigurationID=906df582-99a6-4c42-95e0-a13f220cd505
set UpdateChannel=PerpetualVL2019
set ProductID=ProPlus2019Volume
set ProductKey=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
set OfficeVersion=Office 2019
goto ConfigGen

@rem Генерация XML конфигурационного файла
:ConfigGen
cls
echo ^|
echo ^|   ╠══╦══════════════════════════════════════════╦══╣
echo ^|      ║ Генерация XML конфигурационного файла... ║
echo ^|   ╠══╩══════════════════════════════════════════╩══╣
echo ^|
(
    echo ^<Configuration ID="%ConfigurationID%"^>
    echo   ^<Add OfficeClientEdition="%OfficeEdition%" Channel="%UpdateChannel%"^>
    echo     ^<Product ID="%ProductID%" PIDKEY="%ProductKey%"^>
    echo       ^<Language ID="ru-ru" /^>
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
) > %TEMP%\OfficeSetupFiles\Config.xml
echo ^| Конфигурация сохранена в ^> %TEMP%\OfficeSetupFiles\Config.xml
timeout 1 >nul
goto OfficeExtracterDownload

@rem Загрузка Office
:OfficeExtracterDownload
set PATH=%TEMP%\OfficeSetupFiles\
set ExtractorPath=%TEMP%\OfficeSetupFiles\OfficeExtracter.exe
set SetupPath=%TEMP%\OfficeSetupFiles\setup.exe
setlocal
@rem Загрузка файла с помощью curl
cls
echo ^|
echo ^|   ╠══╦═══════════════════════════════════╦══╣
echo ^|      ║ Скачивание OfficeExtracter.exe... ║
echo ^|   ╠══╩═══════════════════════════════════╩══╣
echo ^|
curl -L -s -o %ExtractorPath% https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe
@rem Проверка, был ли файл загружен успешно
if exist %ExtractorPath% (
    echo ^| Успешно скачено в %PATH%
    timeout 1 >nul
) else (
    echo Ошибка при скачивание OfficeExtracter.exe
    timeout 10 >nul && exit
)
endlocal
goto Extracting

:Extracting
cls
echo ^|
echo ^|   ╠══╦═════════════════════════════╦══╣
echo ^|      ║ Распаковка файлов Office... ║
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
echo ^|      ║ Запуск установщика Office... ║
echo ^|   ╠══╩══════════════════════════════╩══╣
echo ^|
echo ^| Установщик для %OfficeVersion% %OfficeEdition%-Бит запущен...
start %SetupPath% /configure %PATH%Config.xml
timeout 2 >nul
echo ^|
echo ^| Спасибо за использование моего скрипта. Пожалуйста, поддержите меня на Ko-fi: https://ko-fi.com/MaximeriX
echo ^| Нажмите 1, чтобы открыть ссылку
echo ^| Нажмите 2, чтобы выйти
choice /C:12 /M "| >" /N
set Donation=%errorlevel%
if %Donation% == 1 start https://ko-fi.com/MaximeriX (
) else ( 
echo ^| Выход... && timeout 2 >nul && exit
)
@endlocal