Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Popup "We are initiating an upgrade process for Office 2013. Please make sure to save all your work. Your session on this device will log out automatically within the next 15 minutes.", 0, "Office Upgrade", 0 + 32

WScript.Quit
