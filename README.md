# Add Application to Windows 10 Taskbar
# Remove Application from Windows 10 Taskbar
# PowerShell-Functions for Taskbar-Pinning in Windows 10

This small powershell Script adds or removes your own Applications from / to Windows 10 Taskbar (Taskbar-Pinning). The functionality is a bit tricky in Windows 10 as Microsoft doesn't like to offer Applications to pin themselves during installation.

Script: [TaskbarPinning.ps1](TaskbarPinning.ps1)

Tested with: Windows 10 v1709 (English & German)

## Usage: Variant 1: Pin / Unpin regular Executables (no UWP Apps)
```powershell
Import-Module .\TaskbarPinning.ps1

# Add Taskbar-Pinning
Add-TaskbarPinning("C:\Windows\Notepad.exe")

# Remove Taskbar-Pinning
Remove-TaskbarPinning("C:\Windows\Notepad.exe")
```

## Usage: Variant 2: Pin / Unpin Applications and UWP Apps already listed in Start Menu
Note: Handling Add-/Remove-TaskbarPinningApp can only be done by Processes named `explorer`
Workaround to do this with Powershell: Make a copy of PowerShell named explorer.exe in Temp-Directory and use this
```powershell
copy $env:windir\System32\WindowsPowerShell\v1.0\powershell.exe $env:TEMP\explorer.exe

& $env:TEMP\explorer.exe
Import-Module .\TaskbarPinning.ps1
```

List all by Explorer known Apps containing a given Subtring
```powershell
List-ExplorerApps("Microsoft")
```
Shows a List like this:
```
Name                                                                                             Value
----                                                                                             -----
{6D809377-6AF0-444B-8957-A3773F02200E}\Microsoft\Web Platform Installer\WebPlatformInstaller.exe Microsoft Web Platform Installer
Microsoft.ConnectivityStore_8wekyb3d8bbwe!App                                                    Microsoft Wi-Fi
Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge                                              Microsoft Edge
Microsoft.MicrosoftSolitaireCollection_8wekyb3d8bbwe!App                                         Microsoft Solitaire Collection
Microsoft.WindowsStore_8wekyb3d8bbwe!App                                                         Microsoft Store
```

Add or Remove TaskbarPinning of an App
```powershell
# Add or Remove UWP Apps like Edge
Add-TaskbarPinningApp("Microsoft Edge")
Remove-TaskbarPinningApp("Microsoft Edge")
```
