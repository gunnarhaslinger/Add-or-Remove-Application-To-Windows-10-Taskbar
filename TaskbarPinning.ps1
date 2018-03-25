# --- Variant 1: Pin / Unpin regular Win32/64 Executables (no UWP Apps) not already found in Start Menu -----
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
   $ExistingTargets -contains $PinningTarget
}

function Add-TaskbarPinning ($PinningTarget) {
     if (Get-TaskbarPinning($PinningTarget)) { Write-Host "Pinning for $PinningTarget already exists!" } 
     else {Toggle-TaskbarPinning($PinningTarget)}
}

function Remove-TaskbarPinning ($PinningTarget) {
     if (Get-TaskbarPinning($PinningTarget)) { Toggle-TaskbarPinning($PinningTarget) } 
     else { Write-Host "Pinning for $PinningTarget does not exist!" }
}


# --- Variant 2: Pin / Unpin Applications and UWP Apps already listed in Start Menu
function getExplorerVerb([string]$verb) {
    $getstring = @'
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

        public static string GetString(uint strId) {
            IntPtr intPtr = GetModuleHandle("shell32.dll");
            StringBuilder sb = new StringBuilder(255);
            LoadString(intPtr, strId, sb, sb.Capacity);
            return sb.ToString();
        }
'@
    $getstring = Add-Type $getstring -PassThru -Name GetStr -Using System.Text

    if ($verb -eq "PinToTaskbar")     { $getstring[0]::GetString(5386) }  # String: Pin to Taskbar
    if ($verb -eq "UnpinFromTaskbar") { $getstring[0]::GetString(5387) }  # String: Unpin from taskbar
    if ($verb -eq "PinToStart")       { $getstring[0]::GetString(51201) } # String: Pin to start
    if ($verb -eq "UnpinFromStart")   { $getstring[0]::GetString(51394) } # String: Unpin from start
}


function Get-ExplorerApps([string]$AppName) {
    $apps = (New-Object -Com Shell.Application).NameSpace("shell:::{4234d49b-0245-4df3-b780-3893943456e1}").Items()
    $apps | Where {$_.Name -like $AppName -or $app.Path -like $AppName}     
}

function List-ExplorerApps() { List-ExplorerApps(""); }
function List-ExplorerApps([string]$AppName) {
    $apps = Get-ExplorerApps("*$AppName*")
    $AppList = @{};
    foreach ($app in $apps) { $AppList.Add($app.Path, $app.Name) }
    $AppList | Format-Table -AutoSize
}

function Configure-TaskbarPinningApp([string]$AppName, [string]$Verb) {
    $myProcessName = Get-Process | where {$_.ID -eq $pid} | % {$_.ProcessName}
    if (-not ($myProcessName -like "explorer")) { Write-Host "ERROR: Configuring PinningApp can only be done by Processes named 'explorer' but current ProcessName is $myProcessName" }

    $apps = Get-ExplorerApps($AppName)
    if ($apps.Count -eq 0) { Write-Host "Error: No App with exact Path or Name '$AppName' found" }
    $ExplorerVerb = getExplorerVerb($Verb);
    foreach ($app in $apps) {
        $done = "False (Verb $Verb not found)"
        $app.Verbs() | Where {$_.Name -eq $ExplorerVerb} | ForEach {$_.DoIt(); $done=$true }
        Write-Host $verb $app.Name "-> Result:" $done
    }
}

function Remove-TaskbarPinningApp([string]$AppName) { Configure-TaskbarPinningApp $AppName "UnpinFromTaskbar" }
function Add-TaskbarPinningApp([string]$AppName)    { Configure-TaskbarPinningApp $AppName "PinToTaskbar" }
   


# --- Examples Variant 1: Pin / Unpin regular Win32/64 Executables (no UWP Apps) not already found in Start Menu -----

## Add Taskbar-Pinning
# Add-TaskbarPinning("C:\Windows\Notepad.exe")

## Remove Taskbar-Pinning
# Remove-TaskbarPinning("C:\Windows\Notepad.exe")


# --- Examples Variant 2: Pin / Unpin Applications and UWP Apps already listed in Start Menu
## List all Explorer known Apps (Name and Path)
# List-ExplorerApps
#
## List all by Explorer known Apps containing a given Subtring
# List-ExplorerApps("Edge")
#
## Add or Remove UWP Apps like Edge
# Add-TaskbarPinningApp("Microsoft Edge")
# Remove-TaskbarPinningApp("Microsoft Edge")
#
# Note: Handling Add-/Remove-TaskbarPinningApp can only be done by Processes named "explorer.exe"
# Workaround to do this with Powershell: Make a copy of PowerShell named explorer.exe in Temp-Directory and use this
# copy $env:windir\System32\WindowsPowerShell\v1.0\powershell.exe $env:TEMP\explorer.exe
