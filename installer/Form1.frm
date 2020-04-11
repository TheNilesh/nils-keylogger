VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "niL Loggersetting"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   2610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRun 
      Caption         =   "&Run At Startup"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Settings"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1575
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
        MsgBox "Error " & Err.Number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
    End If

    ' See if we should run at startup.
    If run_at_startup Then
        ' Create the key.
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, "explorer", 0, REG_SZ, _
            ByVal key_value, Len(key_value))

        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & Err.Number & " setting key" & _
                vbCrLf & Err.Description
        End If
    Else
        ' Delete the value.
        RegDeleteValue hKey, app_name
    End If

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    MsgBox Err.Number & " " & Err.Description
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



Private Sub Command1_Click()
' Clear or set the key that makes the program run at startup.

    If m_IgnoreEvents Then Exit Sub

    SetRunAtStartup "explorer", "c:\windows\system32", _
        (chkRun.Value = vbChecked)
        
        MsgBox "Settings saved!", vbOKOnly, "Settings"
        End
End Sub

Private Sub Form_Load()
 ' See if the program is set to run at startup.
    m_IgnoreEvents = True
    Dim keycheck As Boolean
    If WillRunAtStartup("explorer") Then
        chkRun.Value = vbChecked
        keycheck = True
    Else
        chkRun.Value = vbUnchecked: keycheck = False
    End If
    m_IgnoreEvents = False
End Sub
