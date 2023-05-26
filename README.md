# Close-Windows-Popup
Visual Basic Script to close (ALT+F4) anoying popup window in MS Windows.

## Configuring
Replace `YourApp` with the name of the application/Window you want to close.
https://github.com/TBuryanek/Close-Windows-Popup/blob/043cc369f22e22810e8ade29fcb633af5983794e/src/close_popup.vbs#L4

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
