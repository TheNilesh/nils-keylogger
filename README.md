# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2010 June 25**

--------------------------

- Installs at `C:\WINDOWS\explorer.exe`
- Stores keylogs in `c:\Windows\system32\sysResource\<incrID>.DAT`
- Also stores active window title
- Installer takes care of HKCU startup entry only.
- Introduced log reader for improving raw logs, replace [ENTER] with new line etc.