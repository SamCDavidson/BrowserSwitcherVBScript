 Msgbox " PRESS OK and temporarily set Chrome as your default browser.",vbExclamation+vbSystemModal,"Step 1 of 4"

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "%windir%\system32\control.exe /name Microsoft.DefaultPrograms /page pageDefaultProgram\pageAdvancedSettings?pszAppName=Google%20Chrome"
WScript.Sleep 2000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys " "
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WScript.Sleep 2000
WshShell.SendKeys " "
Msgbox "Your default browser is now Google Chrome.      PRESS OK and launch Chrome !",vbInformation+vbSystemModal,"Step 2 of 4"

Set WSHShell = CreateObject("WScript.Shell")
result = WSHShell.Run("c:\PROGRA~2\google\chrome\application\chrome.exe --restore-last-session",,True)
msgbox "When you are finished with Google Chrome, PRESS OK and SWITCH BACK TO INTERNET EXPLORER.   It is OK to move this popup box out of the way, it remains open and on top of your desktop until you're finished.",vbInformation+vbSystemModal,"Step 3 of 4"

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "%systemroot%\system32\control.exe /name Microsoft.DefaultPrograms /page pageDefaultProgram\pageAdvancedSettings?pszAppName=Internet%20Explorer"
WScript.Sleep 2000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys " "
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WScript.Sleep 2000
WshShell.SendKeys " "
Msgbox "Thank you for using Chrome! Your Default web browser is back to Microsoft Internet Explorer.      PRESS OK to EXIT.",vbInformation+vbSystemModal,"Step 4 of FINISHED !"