VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Please Register - niL's KeyLogger3"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   Icon            =   "Register.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtTry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdTry 
         Caption         =   "&Continue."
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "&Register !"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtActCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtgCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   810
         Left            =   960
         Picture         =   "Register.frx":57E2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4155
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   120
         Picture         =   "Register.frx":6A13
         Stretch         =   -1  'True
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://niLsKeyLogger.blogspot.com"
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
         Left            =   0
         TabIndex        =   9
         Top             =   2760
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "License Key:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Activation Key :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblTry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remaining Try :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   450
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Get system directory part
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long
Dim SysDir As String
Dim sDr As String                   'Get system drive
Private Sub cmdReg_Click()
If UCase(txtActCode) = genKey(txtgCode) Then
    Call MakeRegistered("C:\Documents and Settings\All Users\Application Data\InstallShield\UpdateService\Q3FD.GML")
    MsgBox "Thank You for purchasing niL's KeyLogger!"
    Unload Me: frmMain.Show
Else
    MsgBox "                    Activation Key Incorrect!" & vbNewLine & "Buy a Activation Key from http://niLsKeyLogger.blogspot.com", vbCritical, "Incorrect Code"
End If

End Sub

Private Sub cmdTry_Click()
   frmMain.Show: Me.Hide
End Sub

Private Sub Form_Load()

Call RegOCX

'******************Activate/Trial Part
If Dir("C:\Documents and Settings\All Users\Application Data\Micro\explorer.exe") <> "" Then  'KLG is installed

        Randomize
        txtgCode = CurrentCode("C:\Documents and Settings\All Users\Application Data\InstallShield\UpdateService\Q3FD.GML")
    If txtgCode <> "Registered" Then
        txtTry = Val(RemainingTry("C:\Documents and Settings\All Users\Application Data\InstallShield\UpdateService\Q3FRT.GLT")) + 1
        If Val(txtTry) < 0 Then txtTry.Visible = False: lblTry.Caption = "Your Free Trial have been expired.": lblTry.ForeColor = &HFF&
         frmMain.txtRegDet.Text = "Licence Key : " & txtgCode.Text: Exit Sub
    Else
     frmMain.txtRegDet.Text = "Registered LogReader 3": Me.Hide: frmMain.Show
    End If
Else
    frmMain.txtRegDet.Text = "Portable LogReader 3": frmMain.Show: Unload Me
End If

End Sub

'*****************************************Function for Activation Key and Trial remainig*************

Private Function CurrentCode(Filepath As String) As String
If Dir(Filepath) <> "" Then
    Open Filepath For Input As 2
    Input #2, CurrentCode
    Close #2
Else
    CurrentCode = GenerateCode(12)
    Open Filepath For Output As 2
    Write #2, CurrentCode
    Close #2
End If
End Function
Private Function RemainingTry(Filepath As String) As Integer
If Dir(Filepath) <> "" Then
    Open Filepath For Input As 2
    Input #2, RemainingTry
    Close #2
End If
End Function
Private Function GenerateCode(CodeLength As Integer) As String
Dim i As Integer
For i = 1 To CodeLength
GenerateCode = GenerateCode & BringChar(Int(Rnd * 36))
Next i

End Function

Private Function genKey(fromThis As String) As String
Dim i As Integer
For i = 1 To Len(fromThis)
    If i = 3 Or i = 7 Then
        genKey = genKey & BringChar(Asc(Mid(fromThis, i, 1)) * 3 Mod 36)
    ElseIf Mid(fromThis, i, 1) = 8 Then
        genKey = genKey
    Else
        genKey = genKey & BringChar(Asc(Mid(fromThis, i, 1)) * 7 Mod 36)
    End If
Next i

End Function
Private Function BringChar(cCode As Integer) As String
Select Case cCode
Case Is < 10
    BringChar = Chr(48 + cCode)
Case Is > 9
    BringChar = Chr(55 + cCode)
End Select
End Function

Private Sub MakeRegistered(Filepath As String)
    Open Filepath For Output As 2
    Write #2, "Registered"
    Close #2
End Sub

Private Sub RegOCX_old()
'Call loadUnamesysDir

If Dir("C:\Documents and settings\All Users\Application Data\Microsoft\COMDLG32.OCX") <> "" And Dir("C:\Windows\system32\COMDLG32.OCX") <> "" Then 'Not Installed COMDLG32.OCX
MsgBox "not C"
FileCopy App.Path & "\file1.CAB", "C:\Documents and settings\All Users\Application Data\Microsoft\COMDLG32.OCX"
Shell "regsvr32 /s" & "C:\Documents and settings\All Users\Application Data\COMDLG32.OCX"
End If

If Dir("C:\Documents and settings\All Users\Application Data\MSCOMCT2.OCX") <> "" And Dir("C:\Windows\system32\MSCOMCT2.OCX") <> "" Then 'not inst MSCOMCT@.OCX
MsgBox "not M"
FileCopy App.Path & "\file2.CAB", "C:\Documents and settings\All Users\Application Data\Microsoft\MSCOMCT2.OCX"
Shell "regsvr32 /s" & "C:\Documents and settings\All Users\Application Data\Microsoft\MSCOMCT2.OCX"
End If

On Error Resume Next
Kill App.Path & "\file1.CAB"
Kill App.Path & "\file2.CAB"

End Sub
Private Sub RegOCX()
'Register OCX extracted in Temp folder
On Error GoTo err
If Dir("C:\Windows\system32\COMDLG32.OCX") <> "" Then 'Not Installed COMDLG32.OCX
'MsgBox ("Already Registered COMDLG32.OCX")
Else
'MsgBox ("Not Registered COMDLG32.OCX")
Shell "regsvr32 /s " & App.Path & "\COMDLG32.OCX"
End If

If Dir("C:\Windows\system32\MSCOMCT2.OCX") <> "" Then 'not inst MSCOMCT2.OCX
'MsgBox ("Already Registered MSCOMCT2.OCX")
Else
Shell "regsvr32 /s " & App.Path & "\MSCOMCT2.OCX"
'MsgBox ("Not Registered MSCOMCT2.OCX")
End If

err:
If err.Number <> 0 Then MsgBox "Administrator on this computer must run this App once.": End
End Sub

