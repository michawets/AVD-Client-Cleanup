param (
    [switch]$SilentRun,
    [switch]$RemovePersonalSettings
)

try {
    Stop-Transcript
} 
catch {}

$ErrorActionPreference = "Stop"
$ScriptLocation = "c:\AvdClientRepair"

if ($null -eq (Get-Item -Path $ScriptLocation -ErrorAction SilentlyContinue)) {
    $dummy = New-Item -Path $ScriptLocation -ItemType Directory -Force
}
$logFile = ("{0}\CleanupAvdClient_{1}.log" -f $ScriptLocation, ((Get-Date).ToString("o").Replace(":", "_")))
Start-Transcript -Path $logFile
Write-Host ("Started at {0}" -f (Get-Date).ToString("o")) -ForegroundColor Cyan


#region functions
function Get-PromptResponse {
    param (
        [string]$Prompt
    )
    if (!($SilentRun)) {
        do {
            $myResponse = Read-Host -Prompt $Prompt
        } while ($myResponse -ne "y" -AND $myResponse -ne "n")
    }
    else {
        $myResponse = "Y"
    }
    return $myResponse
}
#endregion

Write-Host ("Backups can be found here: '{0}'" -f $ScriptLocation) -ForegroundColor Cyan

#Move older backups from previous runs (if there)
$tempFolderName = ("backup_{0}" -f (Get-Date).ToString("o").Replace(":", "_"))
$tempFolderPath = [System.IO.Path]::Combine($ScriptLocation, $tempFolderName)
if ($null -eq (Get-Item -Path $tempFolderPath -ErrorAction SilentlyContinue)) {
    $dummy = New-Item -Path $tempFolderPath -ItemType Directory -Force
}
Move-Item -Path ("{0}\*.reg" -f $ScriptLocation) -Destination $tempFolderPath

#Stopping process if there
Write-Host "Stopping running instances of Remote Desktop client (if running)"
Get-Process -Name "msrdcw" -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process -Name "msrdc" -ErrorAction SilentlyContinue | Stop-Process -Force

#region User Installation
Write-Host "User installation" -ForegroundColor Cyan
#Get list of installed location IDs
$installedGuidsPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Installer\UpgradeCodes\15E4D53F33F0C6248B8AE3FCA4BDCFF5"
Write-Host ("Testing Installed Guid Path in Registry '{0}' (User installation)" -f $installedGuidsPath)
if (!(Test-Path -Path $installedGuidsPath)) {
    Write-Host "Installed Guid Path in Registry not found. Skipping User installation path"
}
else {
    Write-Host "OK" -ForegroundColor Green

    #Backup
    Write-Host ("Creating backup of Registry '{0}'" -f $installedGuidsPath)
    & reg export "HKEY_CURRENT_USER\Software\Microsoft\Installer\UpgradeCodes\15E4D53F33F0C6248B8AE3FCA4BDCFF5" "c:\AvdClientRepair\15E4D53F33F0C6248B8AE3FCA4BDCFF5_UserInstall.reg"
    Write-Host "OK" -ForegroundColor Green

    #get Guids
    Write-Host ("Getting installed GUIDs...")
    $listOfInstalledGuids = Get-Item -Path $installedGuidsPath
    Write-Host ("Found {0} Installed GUID(s). Checking all if any found..." -f $listOfInstalledGuids.Property.Count)
    
    #Loop
    foreach ($installedGuid in $listOfInstalledGuids.Property) {
        #test if path is still valid (can be older install location)
        $productPath = ("Registry::HKEY_CURRENT_USER\Software\Microsoft\Installer\Products\{0}" -f $installedGuid)
        if (!(Test-Path -Path $productPath)) {
            Write-Host ("Trace found of old installation! ==> '{0}'" -f $productPath) -ForegroundColor Cyan

            #Removing
            Write-Host ("Do you want to remove this old entry? (Recommended to remove!)") -ForegroundColor Cyan
            $response = Get-PromptResponse -Prompt "Y/N"
            if ($response -eq "Y") {
                Remove-ItemProperty -Path $installedGuidsPath -Name $installedGuid -Force
                Write-Host "Removed" -ForegroundColor Green
            }
            else {
                Write-Host "Skipping..."
            }
        }
        else {
            Write-Host ("Active installation found at location: '{0}'" -f $productPath) -ForegroundColor Green
            #Backup
            $shortProductPath = $productPath.Replace("Registry::", "")
            Write-Host ("Creating backup of Registry '{0}'" -f $shortProductPath)
            & reg export $shortProductPath "c:\AvdClientRepair\$installedGuid-userinstall.reg"
            Write-Host "OK" -ForegroundColor Green

            Write-Host ("Getting Uninstall GUID from Product...")
            #get Details
            $ProductDetails = Get-Item -Path $productPath
            #Check ProductIcon for GUID
            if ($null -eq $ProductDetails.GetValue("ProductIcon")) {
                Write-Warning "BUG! ProductIcon not found?? Please report to admin"
                continue
            }
    
            #Get ProductGuid
            $ProductId = $ProductDetails.GetValue("ProductIcon").Split("\") | Where-Object { $_ -like "{*" }
            if ($null -eq $ProductId) {
                Write-Warning "BUG! ProductId not found?? Please report to admin"
                continue
            }
            Write-Host "OK" -ForegroundColor Green

            #test if uninstall path is found
            $uninstallRegPath = ("Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{0}" -f $ProductId)
            Write-Host ("Testing Installed Path in Registry '{0}'" -f $uninstallRegPath)
            if (!(Test-Path -Path $uninstallRegPath)) {
                Write-Warning ("Installed Path not found @ '{0}'. Possible old intallation trace..." -f $uninstallRegPath)
                #Removing
                Write-Host ("Do you want to remove this old entry? (Recommended to remove!)") -ForegroundColor Cyan
                $response = Get-PromptResponse -Prompt "Y/N"
                if ($response -eq "Y") {
                    Remove-Item -Path $productPath -Recurse -Force
                    Remove-ItemProperty -Path $installedGuidsPath -Name $installedGuid -Force
                    Write-Host "Removed" -ForegroundColor Green
                }
                else {
                    Write-Host "Skipping..."
                }
                continue
            }
            Write-Host "OK" -ForegroundColor Green

            Write-Host "Validating content..."
            $UninstallInfo = Get-Item -Path $uninstallRegPath
            #test if uninstall path is from AVD client
            $installPath = $UninstallInfo.GetValue("InstallLocation")
            if (!($installPath -like "*\Remote Desktop\*")) {
                Write-Warning ("BUG! InstallLocation is not for AVD?? ('{0}') Please report to admin" -f $installPath)
                continue
            }

            Write-Host ("Installed Path in Registry is OK! '{0}'" -f $installPath) -ForegroundColor Green
            #Backup
            $shortInstallPath = $uninstallRegPath.Replace("Registry::", "")
            Write-Host ("Creating backup of Registry '{0}'" -f $shortInstallPath)
            & reg export $shortInstallPath "c:\AvdClientRepair\$ProductId.reg"
            Write-Host "OK" -ForegroundColor Green

            Write-Host ("Do you want to remove this Uninstall Registry entry? (Recommended to remove!)") -ForegroundColor Cyan
            $response = Get-PromptResponse -Prompt "Y/N"
            if ($response -eq "Y") {
                Remove-Item -Path $uninstallRegPath -Force
                Write-Host "Removed" -ForegroundColor Green
            }
            else {
                Write-Host "Skipping (NOT RECOMMENDED AND NO GOOD CLEANUP IS DONE)..." -ForegroundColor Red
                continue
            }

            #Check physical path of AVD client
            Write-Host "Checking Installation Path of AVD client..."
            if (!(Test-Path -Path $installPath)) {
                Write-Warning ("BUG! InstallLocation is not found ('{0}') Please report to admin" -f $installPath)
                continue
            }

            Write-Host ("Installed Path is OK! '{0}'" -f $installPath) -ForegroundColor Green
            Write-Host ("Do you want to remove this Folder? (Recommended to remove!)") -ForegroundColor Cyan
            $response = Get-PromptResponse -Prompt "Y/N"
            if ($response -eq "Y") {
                Remove-Item -Path $installPath -Recurse -Force
                Write-Host "Removed" -ForegroundColor Green
            }
            else {
                Write-Host "Skipping (NOT RECOMMENDED AND NO GOOD CLEANUP IS DONE)..." -ForegroundColor Red
                continue
            }

            #Cleanup of the Product Path (final step)
            Remove-ItemProperty -Path $installedGuidsPath -Name $installedGuid -Force
            Remove-Item -Path $productPath -Recurse -Force
        }
    }
    Write-Host "Loop done" -ForegroundColor Green
}
#endregion

#region uninstall cleanup (leftovers)
#Loop through the uninstall list
Write-Host "Looking in the Uninstall Registry keys..."
$uninstallRootRegPath = ("Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
$uninstallRegList = Get-ChildItem -Path $uninstallRootRegPath
foreach ($uninstallRegItem in $uninstallRegList) {
    if (!($uninstallRegItem.Property -contains "InstallLocation")) {
        continue
    }
    $installPath = $uninstallRegItem.GetValue("InstallLocation")
    if (!($installPath -like "*\Remote Desktop\*")) {
        continue
    }
    
    Write-Host ("Found an Install Path in Registry: '{0}'" -f $uninstallRegItem.Name) -ForegroundColor Green
    #Backup
    Write-Host ("Creating backup of Registry '{0}'" -f $uninstallRegItem.Name)
    & reg export $uninstallRegItem.Name ("c:\AvdClientRepair\{0}.reg" -f $uninstallRegItem.PSChildName)
    Write-Host "OK" -ForegroundColor Green

    #Check physical path of AVD client
    Write-Host "Checking Installation Path of AVD client..."
    if (!(Test-Path -Path $installPath)) {
        Write-Host ("InstallLocation is not found ('{0}'). Removed already during Machine/User removal." -f $installPath)
        Remove-Item -Path ("Registry::{0}" -f $uninstallRegItem.Name) -Recurse -Force
        continue
    }

    Write-Host ("Installed Path is OK! '{0}'" -f $installPath) -ForegroundColor Green
    Write-Host ("Do you want to remove this Folder? (Recommended to remove!)") -ForegroundColor Cyan
    $response = Get-PromptResponse -Prompt "Y/N"
    if ($response -eq "Y") {
        Remove-Item -Path $installPath -Recurse -Force
        Write-Host "Removed" -ForegroundColor Green
    }
    else {
        Write-Host "Skipping (NOT RECOMMENDED AND NO GOOD CLEANUP IS DONE)..." -ForegroundColor Red
        continue
    }

    #Cleanup of the Product Path (final step)
    Remove-Item -Path ("Registry::{0}" -f $uninstallRegItem.Name) -Recurse -Force
}
#endregion

#region Personal Settings
if ($RemovePersonalSettings) {
    $personalRegPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\RdClientRadc"
    Write-Host ("Checking Personal Registry Path ('{0}')" -f $personalRegPath)
    if (!(Test-Path -Path $personalRegPath)) {
        Write-Warning ("BUG! Personal Registry Path is not found ('{0}') Please report to admin" -f $personalRegPath)
    }
    else {
        Write-Host "OK" -ForegroundColor Green

        #Backup
        $shortPersonalPath = $personalRegPath.Replace("Registry::", "")
        Write-Host ("Creating backup of Registry '{0}'" -f $shortPersonalPath)
        & reg export $shortPersonalPath "c:\AvdClientRepair\RdClientRadc.reg"
        Write-Host "OK" -ForegroundColor Green

        Write-Host ("Do you want to remove this Personal Registry entry? (Recommended to remove!)") -ForegroundColor Cyan
        $response = Get-PromptResponse -Prompt "Y/N"
        if ($response -eq "Y") {
            Remove-Item -Path $personalRegPath -Recurse -Force
            Write-Host "Removed" -ForegroundColor Green
        }
        else {
            Write-Host "Skipping (NOT RECOMMENDED AND NO GOOD CLEANUP IS DONE)..." -ForegroundColor Red
            continue
        }
    }
}
#endregion

#region StartMenu
Write-Host "Cleaning up Start Menu icons..."
$WScriptShell = New-Object -ComObject WScript.Shell
$UserInstallationStartMenu = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs")
if (Test-Path -Path $UserInstallationStartMenu) {
    Write-Host ("Checking if User Start Menu contains shortcut(s) ('{0}')" -f $UserInstallationStartMenu)
    $defaultUserInstallationStartMenuItemLocation = [System.IO.Path]::Combine($UserInstallationStartMenu, "Remote Desktop.lnk")
    if (Test-Path -Path $defaultUserInstallationStartMenuItemLocation) {
        $defaultUserInstallationStartMenuItem = $WScriptShell.CreateShortcut($defaultUserInstallationStartMenuItemLocation)
        if ($defaultUserInstallationStartMenuItem.TargetPath -like "*msrdcw.exe") {
            Write-Host ("Shortcut found: '{0}'" -f $defaultUserInstallationStartMenuItemLocation) -ForegroundColor Green
            Write-Host ("Do you want to remove this shortcut? (Recommended to remove!)") -ForegroundColor Cyan
            $response = Get-PromptResponse -Prompt "Y/N"
            if ($response -eq "Y") {
                Remove-Item -Path $defaultUserInstallationStartMenuItemLocation -Force
                Write-Host "Removed" -ForegroundColor Green
            }
            else {
                Write-Host "Skipping (NOT RECOMMENDED AND NO GOOD CLEANUP IS DONE)..." -ForegroundColor Red
            }
        }
    }
    $oldUsersRDShortcutFolders = Get-ChildItem -Path $UserInstallationStartMenu -Directory -Filter "*(RD)*"
    if ($null -ne $oldUsersRDShortcutFolders) {
        Write-Host "Found some old shortcut folders in the User Startmenu, this is the list:" -ForegroundColor Green
        $oldUsersRDShortcutFolders | Select-Object Name, LastWriteTime | Format-Table
        Write-Host ("Do you want to remove these shortcut(s)? (Recommended to remove!)") -ForegroundColor Cyan
        $response = Get-PromptResponse -Prompt "Y/N"
        if ($response -eq "Y") {
            $oldUsersRDShortcutFolders | Select-Object FullName -ExpandProperty FullName | Remove-Item -Recurse -Force
            Write-Host "Removed" -ForegroundColor Green
        }
        else {
            Write-Host "Skipping (NOT RECOMMENDED AND NO GOOD CLEANUP IS DONE)..." -ForegroundColor Red
        }
    }
}
#endregion

#Starting to download latest AVD client
Write-Host "Starting to download latest AVD client..."  -ForegroundColor Green
$url = "https://go.microsoft.com/fwlink/?linkid=2068602"
$downloadLocation = [System.IO.Path]::Combine($ScriptLocation, "RemoteDesktop_x64.msi")
Write-Host "Downloading, please wait..."
$download = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Get -OutFile $downloadLocation
Write-Host "Download completed"

Write-Host "Starting installation" -ForegroundColor Green
$installResult = Start-Process -FilePath $downloadLocation -ArgumentList @("/qn", "ALLUSERS=2", "MSIINSTALLPERUSER=1") -Wait -PassThru
if ($installResult.ExitCode -ne 0) {
    Write-Host "WARNING: Error during installation. Please check logfiles" -ForegroundColor Red
    Write-Host "Generating log..."
    Write-Host ($installResult | ConvertTo-Json -Depth 2 -Compress)
    return
}

Write-Host "All Done..." -ForegroundColor Green
Stop-Transcript