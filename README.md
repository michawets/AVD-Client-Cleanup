# AVD-Client-Cleanup

This script is to help you clean up an AVD Client installation.<br/>
This could be helpful when you are unable to uninstall or update an existing installation.<br/>
After the cleanup, your default browser will open to download the latest version of the AVD Client.

 > This script only deals with the Windows Installer edition of the AVD client. The Microsoft Store version is not in scope!

## Parameters

 - -SilentRun<br/>This parameter will remove all prompts and will remove all found entries automatically.
 - -RemovePersonalSettings<br/>This parameter will allow the script to clean up all personal settings in the registry. By default, this will be skipped, keeping your entries in the client when installing the AVD client again.

## Download script

This will help you download the script:

```
New-Item -ItemType Directory -Path "C:\Temp" -Force
Set-Location -Path "C:\Temp"
$Uri = "https://raw.githubusercontent.com/michawets/AVD-Client-Cleanup/main/AVD_Client_cleanup.ps1"
# Download the script
Invoke-WebRequest -Uri $Uri -OutFile ".\AVD_Client_cleanup.ps1"
```

## Example commands

This example will clean up everything found without prompting for confirmation

```
& '.\AVD_Client_cleanup.ps1' -SilentRun
```

This example will clean up without confirmation, and will also clean up Personal Settings from the AVD Client (Full cleanup)

```
& '.\AVD_Client_cleanup.ps1' -SilentRun -RemovePersonalSettings
```

## Cleanup locations

This script will clean up these locations:

* Machine installation Registry location(s)
* User installation Registry location(s)
* Machine installation folder(s) 
* User installation folder(s)
* Start Menu icons for the user
* Start Menu icons on the machine
* Leftover shortcuts pointing to applications/desktops


<br>
<br>
<br>

 > **Legal Disclaimer** 

 > THESE SCRIPTS AND EXAMPLE FILES ARE PROVIDED "AS IS" AND WITHOUT ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, WITHOUT LIMITATION, THE IMPLIED WARRANTIES OF MERCHANTIBILITY AND FITNESS FOR A PARTICULAR PURPOSE.

 > UNDER NO CIRCUMSTANCES SHALL THE AUTHOR BE LIABLE TO YOU OR ANY OTHER PERSON FOR ANY INDIRECT, SPECIAL, INCIDENTAL, OR CONSEQUENTIAL DAMAGES OF ANY KIND RELATED TO OR ARISING OUT OF YOUR USE OF THE SCRIPTS AND EXAMPLE FILES, EVEN IF THE AUTHOR OR USER HAS BEEN INFORMED OF THE POSSIBILITY OF SUCH DAMAGES.
