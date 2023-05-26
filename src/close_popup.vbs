Set wshShell = CreateObject("WScript.Shell") 

Do 
    ret = wshShell.AppActivate("YourApp") 'Here you need to update your pop-up program name
If ret = True Then 
    wshShell.SendKeys "%{F4}" 'ALT F4    
    Exit do
End If 
WScript.Sleep 100 'milliseconds
Loop
