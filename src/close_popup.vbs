Set wshShell = CreateObject("WScript.Shell") 

Do 
ret = wshShell.AppActivate("YourApp") 
If ret = True Then 
    wshShell.SendKeys "%{F4}" 'ALT F4    
    Exit do
End If 
WScript.Sleep 500 
Loop
