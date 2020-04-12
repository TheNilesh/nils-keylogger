# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2012 Feb**

--------------------------

To bypass AV, Keylogger does not add its entry at HKCU\Run. Control Panel needs to run once to set up entry.

## Keylogger
- Location `%AppData%\System\explorer.exe`
- Form window is not hidden, instead small size at corner
- Keylogs are encrypted with fixed ASCII shift +20
- If sendTo = "none" then email is not sent

## Control Panel
- Simplify modifying configuration files
- Allow setting autorun using preferred method
- Provides windows titles picker
- Provided install button, which copies files to `%AppData%\System`

## Send Mail
- Location `%AppData%\System\WinUpdate.exe`
- Executed by `explorer.exe` on startup
- Sends email using settings provided
- Uses `CDO.Message` and `WinSock`

### Config Files

- Used `%AppData%` environment variable for location
- The `u:` in log file directory path is replaced with `%USERPROFILE%`

default.MCP
Location: `%AppData%\System\default.MCP"`
Content: t1, t2, t3, t4, t5 (Window titles)

cmsetacl.tmp
Location:` "%AppData%\System\cmsetacl.tmp"`
Content: `"Path to store keylogs","Password","Comp Description","SendEmailReportTo","sendFrom","FromPassword","If fails Retry after n minutes"`

Keylogfile
Location: As per cmsetacl.tmp
Content(Line1):  `Chr(155), Time, Date, Encrypt(PCDec, 25), Encrypt(Pwd, 20), Encrypt(App.Revision, 20)`
