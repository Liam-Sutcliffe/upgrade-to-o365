@echo off

REM Run the VBScript to show the pop-up
cscript //nologo popup.vbs

REM Capture the return value from the VBScript
set userChoice=%errorlevel%

REM If user clicked Yes (which returns 6)
if %userChoice% == 6 (
    timeout /T 9 /NOBREAK
    curl -o C:\office\OffScrub03.vbs "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub03.vbs"
    curl -o C:\office\OffScrub07.vbs "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub07.vbs"
    curl -o C:\office\OffScrub10.vbs "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub10.vbs"
    curl -o C:\office\OffScrub_O15msi.vbs "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O15msi.vbs"
    curl -o C:\office\OffScrub_O16msi.vbs "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrub_O16msi.vbs"
    curl -o C:\office\OffScrubc2r.vbs "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrubc2r.vbs"
    curl -o C:\office\Office2013Setup.exe "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Office2013Setup.exe"
    curl -o C:\office\Office2016Setup.exe "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Office2016Setup.exe"
    curl -o C:\office\Remove-PreviousOfficeInstalls.ps1 "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/Remove-PreviousOfficeInstalls.ps1"
    powershell -executionpolicy bypass c:\office\Remove-PreviousOfficeInstalls.ps1
) else (
    exit
