# Add or Remove Application from/to Windows 10 Taskbar (Taskbar-Pinning)
This small powershell Script adds or removes your own Applications from / to Windows 10 Taskbar.

Tested with: Windows 10 v1709 (German)

## Usage
```powershell
Import-Module .\Taskbarpinning.ps1

# Add Taskbar-Pinning
Add-TaskbarPinning("C:\Windows\Notepad.exe")

# Remove Taskbar-Pinning
Remove-TaskbarPinning("C:\Windows\Notepad.exe")
```
