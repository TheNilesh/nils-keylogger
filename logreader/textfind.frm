VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Find Text"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Log browser"
      Height          =   10935
      Left            =   16920
      TabIndex        =   14
      Top             =   720
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   10005
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   4215
      End
      Begin VB.CommandButton cmdsDate 
         Caption         =   "FileAttrDates"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   0
         TabIndex        =   16
         Top             =   9960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdFdate 
         Caption         =   "openEach file Load"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtLogname 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Dont show [Enter]"
      Height          =   375
      Left            =   12720
      TabIndex        =   12
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dont Show Clicks"
      Height          =   375
      Left            =   10920
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindLoad 
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   11760
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Load File"
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindFirst 
      Caption         =   "Find First"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   11760
      Width           =   975
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   12240
      Width           =   975
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   12240
      Width           =   975
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   11760
      Width           =   975
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   12240
      Width           =   2055
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   11760
      Width           =   2055
   End
   Begin VB.TextBox txtFile 
      Height          =   11055
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   16695
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Replace with"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   12240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Find substring"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   11760
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Written by Igor Ostrovsky (igor@ostrosoft.com)
'Visual Basic 911 (http://www.ostrosoft.com/vb)
'==============================================
Option Explicit

Private Sub Check1_Click()
If Check1.Value = vbChecked Then Call ReplaceText("[CLICK]", vbCrLf)
If Check1.Value = vbUnchecked Then Call openlog(txtLogname)

End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then Call ReplaceText("[ENTER]", vbCrLf)
If Check2.Value = vbUnchecked Then Call openlog(txtLogname)
End Sub




Private Sub cmdFile_Click()
'dlg.FileName = "*.txt"
dlg.ShowOpen
 Call openlog(dlg.FileName)
End Sub

Private Sub cmdFindFirst_Click()
  If txtFind <> "" Then
    If InStr(txtFile, txtFind) <> 0 Then
      txtFile.SelStart = InStr(txtFile, txtFind) - 1
      txtFile.SelLength = Len(txtFind)
    Else
      MsgBox "Not found"
    End If
  End If
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

Private Sub cmdReplace_Click()
  txtFile.SelText = txtReplace
  cmdFindNext_Click
End Sub

Private Sub cmdReplaceAll_Click()
Call ReplaceText(txtFind, txtReplace)
End Sub
Private Sub ReplaceText(replacethis, withthis As String)
   If replacethis <> "" And replacethis <> withthis Then
    While InStr(txtFile, replacethis) <> 0
      txtFile = Left(txtFile, InStr(txtFile, replacethis) - 1) & withthis & Mid(txtFile, InStr(txtFile, replacethis) + Len(replacethis))
    Wend
  End If
End Sub
Private Sub openlog(logname As String)
 Dim strTemp As String
  txtFile = ""
On Error GoTo err
  If Dir(logname) <> "" Then
  txtLogname = logname
    Open logname For Input As 1
    While Not EOF(1)
      Line Input #1, strTemp
        txtFile = txtFile & strTemp & vbCrLf
    Wend
    Close #1
  Else
    MsgBox "File not found", vbCritical
  End If
  
err:
If err <> 0 And err <> 52 Then MsgBox err.Description
End Sub


Private Sub cmdsDate_Click()
File1.Path = "C:\sysResource"
Dim i As Integer
List1.Clear
For i = 1 To File1.ListCount
  File1.ListIndex = i - 1
List1.AddItem FileDateTime(File1.Path & "\" & File1.FileName)
Next i
End Sub

Private Sub cmdFDate_Click()
File1.Path = "C:\sysResource"

List1.Clear
Dim n As Integer, ldate, ltime, luname As String
For n = 1 To File1.ListCount
Open "C:\sysResource\browse" & n & "xcz.dll" For Input As #1
Input #1, ldate, ltime, luname
List1.AddItem ldate & ltime & luname
Close #1
Next n

End Sub





Private Sub List1_DblClick()
Call openlog("C:\sysResource\browse" & List1.ListIndex + 1 & "xcz.dll")
End Sub

Private Sub txtFind_Change()
If txtFind <> "" Then cmdFindLoad.Visible = True: cmdFindLoad.Caption = "Show Lines containing: " & txtFind
If Trim(txtFind) = "" Then cmdFindLoad.Visible = False
End Sub
