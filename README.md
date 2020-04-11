# nils-keylogger
Keylogger which can record keystrokes typed under specific window titles

Version as on **2010 June 03**

--------------------------

- Uses dynamically linked `KbLog32.dll`
- Stores keylogs in `c:\Windows\system32\sysResource\<incrID>.DAT`
- Also stores active window title
- Password Protected Installer nilesh/password which takes care of HKCU startup entry, copying DLL and EXE etc.