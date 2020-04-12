# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2011 July 07**

--------------------------

TRIAL/EVALUATION Version - Minor UX changes

### Keylogger
- Installs at locaton from where its run first time, means it just adds RunOnce entry for the current path
- Stores keylogs in file as per `LogMode` and `LogDir`
- Stores active window title as well as child window controlled by `UseChildTitle`
- Log first line contains metadata `Time, encCode, App.Revision, username`
- Keylogs are now encrypted with the Caesar cipher. Each character is ASCII shifted by `EncCode`.
- Use INI file `C:\Program Files\Common Files\setting.ini` for settings
- Use INI file using `kernel32 GetPrivateProfileString`, not any OCX control
- INI file path is changed `SPYXX.INI`

### LogReader
- Using native INI function instead of OCX
- INI Setting can be set using nice UI

## Installer
- Provides ability to install/uninstall keylogger