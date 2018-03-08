# Add Application to Windows 10 Taskbar
# Remove Application from Windows 10 Taskbar
# PowerShell-Functions for Taskbar-Pinning in Windows 10

This small powershell Script adds or removes your own Applications from / to Windows 10 Taskbar (Taskbar-Pinning). The functionality is a bit tricky in Windows 10 as Microsoft doesn't like to offer Applications to pin themselves during installation.

Tested with: Windows 10 v1709 (German)

## Usage
```powershell
Import-Module .\Taskbarpinning.ps1

# Add Taskbar-Pinning
Add-TaskbarPinning("C:\Windows\Notepad.exe")

# Remove Taskbar-Pinning
Remove-TaskbarPinning("C:\Windows\Notepad.exe")
```
