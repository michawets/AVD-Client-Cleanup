# Howto

These scripts can help you cleanup a (corrupt) System install & perform a clean User install.

The idea is to

1. Run the sysem cleanup as a Administrator account.
2. Run the user cleanup as the user which will also install the latest client.

The scripts will

* cleanup old installations
* fix settings for Auto Update
* download the latest client
* install the client as a user install

## The script to run as an Administrator

```powershell
New-Item -ItemType Directory -Path "C:\Temp" -Force
Set-Location -Path "C:\Temp"
$Uri = "https://raw.githubusercontent.com/michawets/AVD-Client-Cleanup/main/AVD_Client_cleanup-system.ps1"
# Download the script
Invoke-WebRequest -Uri $Uri -OutFile ".\AVD_Client_cleanup-system.ps1"

& '.\AVD_Client_cleanup-system.ps1' -SilentRun
```

## The script to run as a regular user

```powershell
New-Item -ItemType Directory -Path "C:\Temp" -Force
Set-Location -Path "C:\Temp"
$Uri = "https://raw.githubusercontent.com/michawets/AVD-Client-Cleanup/main/AVD_Client_cleanup-user.ps1"
# Download the script
Invoke-WebRequest -Uri $Uri -OutFile ".\AVD_Client_cleanup-user.ps1"

& '.\AVD_Client_cleanup-user.ps1' -SilentRun
```
