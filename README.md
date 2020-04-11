# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2010 June 26**

--------------------------

### Keylogger
- Installs at `C:\WINDOWS\explorer.exe`
- Stores keylogs in `c:\sysResource\browse<incID>xcz.dll`
- Stores active window title
- Log first line contains metadata `Date,Time,username,Chr(13)`
- Keylogs are now encrypted with the Caesar cipher. Each character is ASCII shifted by 1.
 
### Installer
- No need of installer now!
- Keylogger itself copies itself to system32 directory
- Keylogger itself adds startup registry entry at HKCU RunOnce.

### LogReader
- no change!
