# BatchOfficeInstaller
![BatchOfficeInstaller-Image](https://github.com/user-attachments/assets/142e14be-eaf8-4042-8816-e4e0c5168a31)

**BatchOfficeInstaller** is a simple .bat script for downloading Office **LTSC 2024**, **LTSC 2021**, **2019** and **2016** for **free**.

This script contains **Access**, **Excel**, **OneDrive**, **OneNote**, **Outlook**, **PowerPoint**, **Project**, **Publisher**, **Teams**, **Visio**, and **Word** (Can be changed).
There's no need to download anything manually, the script will handle everything for you.

## System Requirements
- Windows 10 or later
- Server 2019 or later

## Download

Download the [**BatchOfficeInstaller**](https://github.com/MaximeriX/BatchOfficeInstaller/releases/tag/Release-1.0.9) from the releases section.

## How it Works
#### When user selected Office Version and Apps
1. Generates `config.xml`
2. Downloads and runs `officedeploymenttool.exe` from [Microsoft's official site](https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe)
3. Runs `setup.exe /Configure config.xml`
   
*Note*: Every file that the program downloads/generates during office installation is located in `%TEMP%\OfficeSetupFiles`
