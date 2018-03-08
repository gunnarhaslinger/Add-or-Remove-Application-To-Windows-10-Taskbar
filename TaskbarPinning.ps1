Function Toggle-TaskbarPinning($PinningTarget) {
    # Create HKCU:\SOFTWARE\Classes\*\shell\PinMeToTaskBar" with CommandHandler "TaskbarPin" of Explorer
    $Key2 = (Get-Item "HKCU:\SOFTWARE\Classes").OpenSubKey("*", $true)
    $Key3 = $Key2.CreateSubKey("shell", $true)
    $Key4 = $Key3.CreateSubKey("PinMeToTaskBar", $true)

    # Get Explorer Command Handler for TaskbarPinning (currently with v1709 its "{90AA3A4E-1CBA-4233-B8BB-535773D48449}")
    $ValueData = (Get-ItemProperty ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\Windows.taskbarpin")).ExplorerCommandHandler
    $Key4.SetValue("ExplorerCommandHandler", $ValueData)

    # execute Target-Objekts Shell-Verb "PinMeToTaskBar" (which we created just above)
    $Shell = New-Object -ComObject "Shell.Application"
    $Folder = $Shell.Namespace((Get-Item $PinningTarget).DirectoryName)
    $Item = $Folder.ParseName((Get-Item $PinningTarget).Name)
    $Item.InvokeVerb("PinMeToTaskBar")

    # remove the created PinMeToTaskBar entry
    $Key3.DeleteSubKey("PinMeToTaskBar")
    if ($Key3.SubKeyCount -eq 0 -and $Key3.ValueCount -eq 0) { $Key2.DeleteSubKey("shell") }
}

function Get-TaskbarPinningList {
    $Shortcuts = Get-ChildItem -Recurse "$env:USERPROFILE\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Include *.lnk
    $Shell = New-Object -ComObject WScript.Shell
    foreach ($Shortcut in $Shortcuts)
    {
        $Properties = @{
        ShortcutName = $Shortcut.Name;
        ShortcutFull = $Shortcut.FullName;
        ShortcutPath = $shortcut.DirectoryName
        Target = $Shell.CreateShortcut($Shortcut).targetpath
        }
        New-Object PSObject -Property $Properties
    }

    [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null     
}

function Get-TaskbarPinning ($PinningTarget) {
   $ExistingTargets = (Get-TaskbarPinningList).Target
   $ExistingTargets -contains $Target
}

function Add-TaskbarPinning ($PinningTarget) {
     if (Get-TaskbarPinning($PinningTarget)) { Write-Host "Pinning for $PinningTarget already exists!" } 
     else {Change-TaskbarPinning($PinningTarget)}
}

function Remove-TaskbarPinning ($PinningTarget) {
     if (Get-TaskbarPinning($PinningTarget)) { Change-TaskbarPinning($PinningTarget) } 
     else { Write-Host "Pinning for $PinningTarget does not exist!" }
}



# Add Taskbar-Pinning
Set-TaskbarPinning("C:\Windows\Notepad.exe")

# Remove Taskbar-Pinning
# Delete-TaskbarPinning("C:\Windows\Notepad.exe")
