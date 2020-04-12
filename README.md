# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2010 October 01**

--------------------------

### Keylogger
- Installs at locaton from where its run first time, means it just adds RunOnce entry for the current path
- Stores keylogs in file as per `LogMode` and `LogDir`
- Stores active window title as well as child window controlled by `UseChildTitle`
- Log first line contains metadata `Time, encCode, App.Revision, username`
- Keylogs are now encrypted with the Caesar cipher. Each character is ASCII shifted by `EncCode`.
- Use INI file `C:\Program Files\Common Files\setting.ini` for settings
- Use INI file using `kernel32 GetPrivateProfileString`, not any OCX control

### LogReader
- Using native INI function instead of OCX
- INI Setting can be set using nice UI
