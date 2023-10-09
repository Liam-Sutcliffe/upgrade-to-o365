#Run this on the target machine via remote powershell or agent like Atera
New-Item -ItemType Directory -Path "C:\office" -Force
New-Item -ItemType Directory -Path "C:\office\Office365Install" -Force
curl -o C:\office\logoff.ps1 "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/logoff.ps1"
curl -o C:\office\install.ps1 "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/install.ps1"
msg * "We are initiating an upgrade process for Office 2013. Please make sure to save all your work. Your session on this device will log out automatically within the next 15 minutes."
curl -o C:\office\OffScrub03.vbs "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/OffScrub03.vbs"
curl -o C:\office\OffScrub07.vbs "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/OffScrub07.vbs"
curl -o C:\office\OffScrub10.vbs "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/OffScrub10.vbs"
curl -o C:\office\OffScrub_O15msi.vbs "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/OffScrub_O15msi.vbs"
curl -o C:\office\OffScrub_O16msi.vbs "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/OffScrub_O16msi.vbs"
curl -o C:\office\OffScrubc2r.vbs "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/OffScrubc2r.vbs"
curl -o C:\office\Office2013Setup.exe "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/Office2013Setup.exe"
curl -o C:\office\Office2016Setup.exe "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/Office2016Setup.exe"
curl -o C:\office\Remove-PreviousOfficeInstalls.ps1 "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/Remove-PreviousOfficeInstalls.ps1"
curl -o C:\office\ODTSetup.exe "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/ODTSetup.exe"
curl -o C:\office\Office365Install\configuration-Office365-x64.xml "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/main/Office365Install/configuration-Office365-x64.xml"
curl -o C:\office\Office365Install\setup.exe "https://raw.githubusercontent.com/Liam-Sutcliffe/upgrade-to-o365/33d085ed945b324a14543d58e9f2127aae800804/Office365Install/setup.exe"
Start-Sleep -Seconds 180
write-output "3minutes passed"
Start-Sleep -Seconds 180
write-output "6minutes passed"
Start-Sleep -Seconds 180
write-output "9minutes passed"
Start-Sleep -Seconds 180
write-output "12minutes passed"
Start-Sleep -Seconds 180
write-output "15minutes passed"
powershell -executionpolicy bypass c:\office\Logoff.ps1
powershell -executionpolicy bypass c:\office\Remove-PreviousOfficeInstalls.ps1
powershell -executionpolicy bypass c:\office\Install.ps1
Start-Sleep -Seconds 5
Get-ChildItem -Path "C:\office\*" | Remove-Item -Recurse -ErrorAction SilentlyContinue
