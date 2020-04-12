VERSION 5.00
Begin VB.Form frmTitles 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Selected Titles"
   ClientHeight    =   1890
   ClientLeft      =   2775
   ClientTop       =   3645
   ClientWidth     =   5355
   Icon            =   "sLogging.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Choose..."
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtTitle5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   10
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtTitle4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtTitle3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtTitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtTitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Title 5  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Title 4  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Title 3  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Title 2  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title 1  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get system directory part
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long
Dim sDr As String
Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub



Private Sub Command1_Click()

frmMain.fraTitles.Visible = True
frmMain.txtTitle1 = txtTitle1
frmMain.txtTitle2 = txtTitle2
frmMain.txtTitle3 = txtTitle3
frmMain.txtTitle4 = txtTitle4
frmMain.txtTitle5 = txtTitle5
Unload frmOptions
Unload Me

End Sub
Private Sub Form_Load()
Call LoadTitle
End Sub
Private Sub LoadTitle()
'Get sysDirectory
Dim SysDir As String
SysDir = String(80, 0)
Call GetSystemDirectory(SysDir, 80)     'stores global variable sysDir i.e, system32 path.
sDr = Left(SysDir, 2)

Dim t1, t2, t3, t4, t5 As String
On Error GoTo err
Open sDr & "\Documents and Settings\" & frmOptions.Username.Text & "\UserData\titles.dat" For Input As 1
On Error Resume Next
Input #1, t1, t2, t3, t4, t5
Close #1

txtTitle1 = t1
txtTitle2 = t2
txtTitle3 = t3
txtTitle4 = t4
txtTitle5 = t5

err:
If err = 0 Then Exit Sub
If err.Number = 53 Then Call createtitle Else MsgBox err.Description

End Sub
Private Sub createtitle()

Open sDr & "\Documents and Settings\" & frmOptions.Username.Text & "\UserData\titles.dat" For Output As 1
Write #1, "Title1", "Title2", "Title3", "Title4", "Title5"
Close #1
Call LoadTitle
End Sub

Private Sub OKButton_Click()
Open sDr & "\Documents and Settings\" & frmOptions.Username.Text & "\UserData\titles.dat" For Output As 1
Write #1, txtTitle1, txtTitle2, txtTitle3, txtTitle4, txtTitle5
Close #1

Unload Me
End Sub

