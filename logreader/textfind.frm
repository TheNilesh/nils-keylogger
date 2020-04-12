VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Portable Log Reader by itsnilhere@rediffmail.com"
   ClientHeight    =   12300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20235
   Icon            =   "textfind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12300
   ScaleWidth      =   20235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   240
      Top             =   240
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MouseIcon       =   "textfind.frx":57E2
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Picture         =   "textfind.frx":593D
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Update"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Picture         =   "textfind.frx":CE3B
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Open Log"
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Dont Show [TAB]"
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Dont show [Enter]"
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dont Show Clicks"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Log Details:"
      Height          =   1575
      Left            =   16200
      TabIndex        =   6
      Top             =   960
      Width           =   3255
      Begin VB.TextBox txtVersion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtencr 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtShut 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtstart 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "KeyLogger ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Encrypted:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Shutdown at:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Logged on at:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11895
      HideSelection   =   0   'False
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1200
      Width           =   19215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log Tracker:"
      Height          =   855
      Left            =   9000
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdFindLoad 
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Find substring"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtLogname 
      Height          =   225
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "itsnilhere@rediffmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   13560
      TabIndex        =   23
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Check1_Click()
If Check1.Value = vbChecked Then Call ReplaceText("[C]", vbCrLf)
If Check1.Value = vbUnchecked Then Call openlog(txtLogname)

End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then Call ReplaceText("[ENTER]", vbCrLf)
If Check2.Value = vbUnchecked Then Call openlog(txtLogname)
End Sub
Private Sub Check3_Click()
If Check3.Value = vbChecked Then Call ReplaceText("[TAB]", vbTab)
If Check3.Value = vbUnchecked Then Call openlog(txtLogname)
End Sub
Private Sub cmdFindLoad_Click()
If txtLogname = "" Then Exit Sub

  Dim strTemp As String
  txtFile = ""

  If Dir(txtLogname) <> "" Then
    Open txtLogname For Input As 1
    While Not EOF(1)
      Line Input #1, strTemp
    If InStr(LCase(strTemp), LCase(txtFind)) <> 0 Then txtFile = txtFile & strTemp & vbCrLf
    Wend
    Close #1
  Else
    MsgBox "File not found"
  End If

End Sub

Private Sub cmdFindNext_Click()
  If txtFind <> "" Then
    txtFile.SelStart = txtFile.SelStart + 2
    If InStr(txtFile.SelStart, txtFile, txtFind) <> 0 Then
      txtFile.SelStart = InStr(txtFile.SelStart, txtFile, txtFind) - 1
      txtFile.SelLength = Len(txtFind)
    Else
      MsgBox "Not found"
    End If
  End If
End Sub






Private Sub ReplaceText(replacethis, withthis As String)
ProgressBar1.Visible = True
ProgressBar1.Value = 0

If replacethis <> "" And replacethis <> withthis Then
    While InStr(txtFile, replacethis) <> 0
       ProgressBar1.Value = 71
        txtFile = Left(txtFile, InStr(txtFile, replacethis) - 1) & withthis & Mid(txtFile, InStr(txtFile, replacethis) + Len(replacethis))
    Wend
End If

ProgressBar1.Visible = False
End Sub
Private Sub openlog(logname As String)
'get encCode
Close


Dim ltime, luname, version, enccode As String
On Error GoTo err
Open logname For Input As #1
Input #1, ltime, enccode, version, luname
txtUser = luname
txtencr = enccode
txtVersion = version
Close #1
txtShut = FileDateTime(logname)

Me.MousePointer = vbHourglass
ProgressBar1.Value = 0
Dim strTemp As String
txtFile = ""

On Error GoTo err
  If Dir(logname) <> "" Then
  
        Frame3.Visible = False
        ProgressBar1.Visible = True
        ProgressBar1.ToolTipText = "Loading log : " & FileDateTime(logname)
        txtLogname = logname
        
    Open logname For Input As 1
    While Not EOF(1)
         Line Input #1, strTemp
        txtFile = txtFile & decrypt(strTemp) & vbCrLf
         ProgressBar1.Value = (Len(txtFile) / FileLen(logname)) * 100
    Wend
    Close #1
  Else
    MsgBox "File not found", vbCritical
  End If

Me.MousePointer = vbDefault
ProgressBar1.Visible = False
Frame3.Visible = True
err:
If err <> 0 And err <> 52 And err <> 380 Then MsgBox err.Description
If err = 62 Then MsgBox "Sorry, this is not valid Log File.", vbExclamation
Me.MousePointer = vbDefault
ProgressBar1.Visible = False
Frame3.Visible = True
Close
End Sub
Private Function decrypt(iput As String)
Dim i As Integer
For i = 1 To Len(iput)  '
decrypt = decrypt & Chr(Asc(Mid(iput, i, 1)) - txtencr)
Next i
End Function


Private Sub cmdOpen_Click()
'dlg.FileName = "*.txt"
dlg.ShowOpen
 Call openlog(dlg.Filename)
End Sub

Private Sub cmdUpdate_Click()
frmOptions.Show
End Sub


Private Sub Form_Load()
txtFile.Left = 1000
txtFile.Width = Screen.Width - 2000
txtFile.Height = Screen.Height - 2000
txtFile.Top = 1000
Label7.Left = Screen.Width - 6000
End Sub

Private Sub Timer1_Timer()
If Label7.ForeColor = &HFF& Then
Label7.ForeColor = &H8000&
Else: Label7.ForeColor = &HFF&
End If
End Sub

Private Sub txtFind_Change()
If txtFind <> "" Then cmdFindLoad.Visible = True: cmdFindLoad.Caption = "Show Lines containing: " & txtFind
If Trim(txtFind) = "" Then cmdFindLoad.Visible = False
End Sub
