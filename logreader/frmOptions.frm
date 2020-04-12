VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4800
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5535
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSample1 
      Caption         =   "Logging Mode"
      Height          =   2115
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   5295
      Begin VB.OptionButton Option2 
         Caption         =   "Everyday New Logfile"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Only One Logfile."
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "New Logfile on each LogOn. (Recommended)"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Use system Directory"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Log Directory :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Logfile extension :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Logging Setting"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox Check1 
         Caption         =   "Print child window title."
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Remove last character on Backspace."
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         Max             =   100
         SelStart        =   60
         Value           =   60
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   2
         Min             =   -6
         Max             =   6
      End
      Begin VB.Label Label2 
         Caption         =   "Typing speed :"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Encryption strength :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4335
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4335
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1890
      TabIndex        =   0
      Top             =   4335
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
IniWrite "LogSetting", "USEBS", Check2.Value, "c:\Program Files\Common Files\setting.ini"
IniWrite "LogSetting", "UseChildTitle", Check1.Value, "c:\Program Files\Common Files\setting.ini"
IniWrite "LogSetting", "EncCode", Slider1.Value, "c:\Program Files\Common Files\setting.ini"
IniWrite "LogSetting", "extension", Text2.Text, "c:\Program Files\Common Files\setting.ini"
IniWrite "LogSetting", "TimerInt", Slider2.Value, "c:\Program Files\Common Files\setting.ini"
IniWrite "LogSetting", "LogDir", Text1.Text, "c:\Program Files\Common Files\setting.ini"
IniWrite "LogSetting", "LogSysDir", Check3.Value, "c:\Program Files\Common Files\setting.ini"

If Option1.Value = True Then
IniWrite "LogSetting", "LogMode", "0", "c:\Program Files\Common Files\setting.ini"
ElseIf Option2.Value = True Then
IniWrite "LogSetting", "LogMode", "1", "c:\Program Files\Common Files\setting.ini"
ElseIf Option3.Value = True Then
IniWrite "LogSetting", "LogMode", "2", "c:\Program Files\Common Files\setting.ini"
End If

cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MsgBox "Changes will take effect on next Logon."
    Unload Me
End Sub
Private Sub loadsettings()
Check2.Value = INIRead("LogSetting", "USEBS", "c:\Program Files\Common Files\setting.ini")
Check1.Value = INIRead("LogSetting", "UseChildTitle", "c:\Program Files\Common Files\setting.ini")
Slider1.Value = INIRead("LogSetting", "EncCode", "c:\Program Files\Common Files\setting.ini")
Text2.Text = INIRead("LogSetting", "extension", "c:\Program Files\Common Files\setting.ini")
Slider2.Value = INIRead("LogSetting", "TimerInt", "c:\Program Files\Common Files\setting.ini")
Text1.Text = INIRead("LogSetting", "LogDir", "c:\Program Files\Common Files\setting.ini")
Check3.Value = INIRead("LogSetting", "LogSysDir", "c:\Program Files\Common Files\setting.ini")

Dim logmode As Integer
logmode = INIRead("LogSetting", "LogMode", "c:\Program Files\Common Files\setting.ini")
Select Case logmode
Case 0
Option1.Value = True
Case 1
Option2.Value = True
Case 2
Option3.Value = True
End Select
End Sub


Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Call loadsettings
End Sub

Private Sub Slider2_Change()
If Slider2.Value < 10 Then MsgBox "Very Low.": Slider2.Value = 10
End Sub
