# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2012 October**

--------------------------

Used 7-zip to build installer package

## Keylogger
- Location `%ProgramData%\System\explorer.exe`
- Keylogs are NOT encrypted
- If sendTo = "none" then email is not sent

## Installer 7-Zip SFX
- Portable : Can be run from anywhere
- Programmable config.txt with commandline options -ai0, -ai1 etc

## Send Mail
- Location `%ProgramData%\System\wininit.exe`
- Executed by `explorer.exe` on startup
- Sends email in body
- Includes private IP, public IP in email

### Config Files

- Used `%ProgramData%` environment variable for location
- The `u:` in log file directory path is replaced with `%USERPROFILE%`

titles.txt
Location: `%ProgramData%\System\titles.txt"`
Content: One title per line, can be any number of titles

settings.txt
Location: `"%ProgramData%\System\settings.txt"`
Content: `LogDir, CompName, sendTo, sendFrom, Sender Password, RetryTime, PubIPURL`

Keylogfile
Location: LogDir in settings.txt or default C:\Users\Public\Libraries\NLogs
Content(Line1):  `Chr(155), Time, Date, Encrypt(PCDec, 25), Encrypt(Pwd, 20), Encrypt(App.Revision, 20)`
