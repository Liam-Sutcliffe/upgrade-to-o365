Set objShell = WScript.CreateObject("WScript.Shell")
intAnswer = objShell.Popup("We are initiating an upgrade process for Office 2013. Please click 'YES' to proceed with the update. Please make sure to save all your work. Your session on this device will log out automatically within the next 15 minutes", 0, "Office Upgrade", 4 + 32)

WScript.Quit(intAnswer)
