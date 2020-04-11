# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2010 September 30**

--------------------------

### Keylogger
- Installs at locaton from where its run first time, means it just adds RunOnce entry for the current path
- Stores keylogs in file as per `LogMode` and `LogDir`
- Stores active window title as well as child window controlled by `UseChildTitle`
- Log first line contains metadata `Time, encCode, App.Revision, username`
- Keylogs are now encrypted with the Caesar cipher. Each character is ASCII shifted by `EncCode`.
- Use INI file `C:\Program Files\Common Files\setting.ini` for settings. Through `INIEdit.ocx`

#### setting.ini

```
[LogSetting]
USEBS=1
UseChildTitle=0
EncCode=1
extension=.txt
TimerInt=65
LogMode=1
LogDir=sys

; If LogDir=sys then basepath will be system32 dir else LogDir setting value

; If LogMode = 0 Then `<basepath>\sysResource\browse<incrID>xcz<INI.extension>`
; If LogMode = 1 Then `<basepath>\sysResource\browse<ddmm>xcz<INI.extension>`
; If LogMode = 2 Then single log file `<basepath>\sysResource\browsexcz<INI.extension>`

```

### Installer
- No need of installer now!
- Keylogger does not copy itself
- Keylogger itself adds startup registry entry at HKCU RunOnce.

### LogReader
- no change!
