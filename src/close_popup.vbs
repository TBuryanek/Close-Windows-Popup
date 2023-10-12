Set wshShell = CreateObject("WScript.Shell") 

Do 
    ret = wshShell.AppActivate("YourApp") 'Program title name / or PID
    If ret = True Then 
        'Old approach using ALT F4 to close window:
        'wshShell.SendKeys "%{F4}" 'ALT F4 
        'New approach using taskkill:
        wshShell.Run "taskkill /f /im YourApp.exe", , True 'Task/App name
        Exit do
    End If 
    WScript.Sleep 100 'milliseconds
Loop
