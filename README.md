# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2012 Feb**

--------------------------

To bypass AV, Keylogger does not add its entry at HKCU\Run. Control Panel needs to run once to set up entry.

## Keylogger
- Form window is not hidden, instead small size at corner
- Keylogs are encrypted with fixed ASCII shift +20
- If comp Description = "none" then email is not sent

## Control Panel
- Simplify modifying configuration files
- Allow setting autorun using preferred method
- Provides windows titles picker

## Send Mail
- Sends email using settings provided
- Uses `CDO.Message` and `WinSock`

### Config Files

- Order of values in STP.txt changed in favour of keylogger(explorer.exe)

default.MCP
Location: `"C:\Documents and Settings\" & txtUserName & "\Application Data\System\default.MCP"`
Content: t1, t2, t3, t4, t5 (Window titles)

STP.txt
Location:` "C:\Documents and Settings\%USERNAME%\Application Data\System\STP.txt"`
Content: `"Path to store passwords","Password","Comp Description","SendEmailReportTo","sendFrom","FromPassword","If fails Retry after n minutes"`

Keylogfile
Location: As per STP.txt
Content(Line1):  `Chr(155), Time, Date, Encrypt(txtUserName, 25), Encrypt(Pwd/Computer Description, 20), Encrypt(App.Revision, 20)`
