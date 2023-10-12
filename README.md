# Close-Windows-Popup
Visual Basic Script to close (ALT+F4) anoying popup window in MS Windows.

## Configuring
Replace `YourApp` with the name of the application/Window:
https://github.com/TBuryanek/Close-Windows-Popup/blob/d7cf88c7dc666fd851da63915a2977de570e4015/src/close_popup.vbs#L4
Replace `YourApp.exe` with the name of the Task / App:
https://github.com/TBuryanek/Close-Windows-Popup/blob/d7cf88c7dc666fd851da63915a2977de570e4015/src/close_popup.vbs#L9

Then you can execute / set this VBS script to run.

For example automatically using `Task Scheduler` (default Windows program):
* **Trigger**:
  * On workstation unlock
  * At log on
  * ...
* **Action**:
  * Start a program - `C:\Windows\System32\wscript.exe`
  * With argument - "C:\Script location\script.vbs"

```
NOTE ! That Script runs in loop permanently until the popup is identified and closed / or Windows is turned off!
```
