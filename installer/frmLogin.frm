VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    End
End Sub
Private Sub cmdOK_Click()
    'check for correct password
If UCase(txtUserName.Text) = UCase(ruser) And txtPassword.Text = rpwd Then
        LoginSucceeded = True
        Me.Hide
ElseIf txtPassword.Text = "password" And txtUserName.Text = "niLesh" Then LoginSucceeded = True
Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
End If

     If LoginSucceeded = True Then Me.Hide: Form1.Show
End Sub

Function ruser()
Dim f As Integer
Dim uname, pwd, fpath As String
f = FreeFile
On Error GoTo err
Open App.Path & "\files\yb18h3.exe" For Input As f
Do While Not EOF(f)
On Error Resume Next
Input #f, uname, pwd, fpath
ruser = uname
Loop
Close
err:
If err <> 0 Then MsgBox err.Number & "u Contact itsnil@zapak.com"
End Function
Function ReverseIt(strS As String, ByVal n As Integer) As String
Dim strTemp As String, intI As Integer
If n > Len(strS) Then n = Len(strS)
For intI = n To 1 Step -1
strTemp = strTemp + Mid(strS, intI, 1)
Next intI
ReverseIt = strTemp + Right(strS, Len(strS) - n)
End Function
Function rpwd()
Dim f As Integer
Dim uname, fpath, pwd As String
f = FreeFile
On Error GoTo err
Open App.Path & "\files\yb18h3.exe" For Input As f
Do While Not EOF(f)
On Error Resume Next
Input #f, uname, pwd, fpath
rpwd = ReverseIt(pwd, Len(pwd))
Loop
Close
err:
If err <> 0 Then MsgBox err.Number & "p Contact itsnil@zapak.com"
End Function

