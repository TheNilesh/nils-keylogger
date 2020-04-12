VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "niL's KeyLogger  -  LogReader"
   ClientHeight    =   12300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20235
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "textfind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12300
   ScaleWidth      =   20235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRegDet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   600
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   64
      Text            =   "textfind.frx":57E2
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Frame fraTitles 
      BackColor       =   &H80000007&
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   13920
      TabIndex        =   49
      Top             =   2160
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton OKButton 
         Cancel          =   -1  'True
         Caption         =   "&Save"
         Height          =   375
         Left            =   4200
         TabIndex        =   58
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   60
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtTitle1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   50
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtTitle2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   51
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtTitle3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   52
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtTitle4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   54
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtTitle5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   56
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy WindowTitles from Log and Paste into box above."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   840
         TabIndex        =   63
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Titles"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Title 1  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Title 2  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Title 3  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Title 4  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Title 5  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
   End
   Begin VB.TextBox txtUsername 
      Height          =   360
      Left            =   3720
      TabIndex        =   48
      Text            =   "txtUsername"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      TabIndex        =   47
      Top             =   4440
      Width           =   3135
      Begin VB.CheckBox optApply 
         Caption         =   "&Apply"
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Dont Show &Clicks"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Dont show [&Enter]"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Dont Show [&TAB]"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   46
      Top             =   1080
      Width           =   3135
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   16777215
         CalendarTitleBackColor=   12632256
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   4210752
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   20709377
         CurrentDate     =   40653
      End
      Begin VB.Label lblLoadAll 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Load All Logs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame frmChooseLog 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   6720
      TabIndex        =   41
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ListBox List2 
         Height          =   1500
         Left            =   4080
         TabIndex        =   45
         Top             =   4200
         Width           =   1935
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   4080
         TabIndex        =   44
         Top             =   840
         Width           =   2415
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00C0C0C0&
         Height          =   5100
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "All Logs :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Picture         =   "textfind.frx":57ED
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Log In"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   9000
      TabIndex        =   36
      Top             =   4560
      Width           =   3855
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox upwd 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   1680
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "textfind.frx":5AF8
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   38
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Password  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox txtencoded 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      HideSelection   =   0   'False
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   35
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   0
      Top             =   8400
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3960
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MouseIcon       =   "textfind.frx":5F3C
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Picture         =   "textfind.frx":6097
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Setting"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "textfind.frx":730D
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open Log"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Log Details:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   15360
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox txtVersion 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtencr 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtShut 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtstart 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "KeyLogger ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Encrypted:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Session End :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Logged on at:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9615
      HideSelection   =   0   'False
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Top             =   720
      Width           =   15855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   3135
      Begin VB.CommandButton cmdTime 
         Caption         =   "&Go"
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox txtFind 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtHH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton cmdFindNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtMM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label cmdNextSession 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Next Session"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Track Time hh:mm :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Find :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtLogname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   720
      Picture         =   "textfind.frx":8F8A
      Top             =   6840
      Width           =   1800
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://niLsKeyLogger.blogspot.com niLsKeyLogger@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   0
      TabIndex        =   65
      Top             =   8880
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W E L C O M E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3360
      TabIndex        =   39
      Top             =   240
      Width           =   13335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'username import function
Private Declare Function getusername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Get system directory part
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long
Dim sDr As String

Dim unipath As String
Dim i As Integer
Option Explicit
Private Sub OpenLog(LogName As String)

Close

If LogName = "" Then Exit Sub
'get encCode, password
Dim ltime, luname, ldate, version, encCode, noneed, pwd As String
If Dir(LogName) <> "" Then
    On Error GoTo err
    Open LogName For Input As #1
    Input #1, noneed, ltime, encCode, ldate, luname, pwd, version
    Close #1
Else
    MsgBox "Log not available", vbInformation: Exit Sub
End If

If IsNumeric(version) = False Or IsNumeric(encCode) = False Then MsgBox "This is not Valid log File!", vbCritical: Close #1: Exit Sub
txtUser = luname
txtstart = ldate & "  " & ltime
txtencr = encCode
txtVersion = version
txtShut = FileDateTime(LogName)
Label10.Caption = txtstart

If decrypt(pwd, 10) = upwd.Text Or Trim(pwd) = "" Then    'Password is blank or matches
    Call ReadLog(LogName)
Else
    MsgBox "Password Incorrect!", vbCritical
    Frame2.Visible = True
    unipath = LogName
    upwd.SetFocus
    frmChooseLog.Visible = False
End If

err:
If err.Number <> 0 Then MsgBox "This is Invalid Log!", vbCritical

End Sub
Private Sub ReadLog(gLogpath As String)
Frame2.Visible = False
frmChooseLog.Visible = False
Me.MousePointer = vbHourglass
ProgressBar1.Value = 0
Dim strTemp As String
txtFile = ""
Frame3.Visible = False
ProgressBar1.Visible = True
txtLogname = gLogpath

On Error GoTo err
    Open gLogpath For Input As 1
    While Not EOF(1)
         Line Input #1, strTemp
            If Left(strTemp, 2) = Chr(34) & "›" Then 'it founds tag it is first line of log after logon
                  txtFile = txtFile & vbCrLf & "-_-_-_-_NEW SESSION-_-_-_-_-" & vbCrLf
            Else
               txtFile = txtFile & decrypt(strTemp, Val(txtencr)) & vbCrLf
            End If
         ProgressBar1.Value = (Len(txtFile) / FileLen(gLogpath)) * 100
    Wend
    Close #1

Me.MousePointer = vbDefault
ProgressBar1.Visible = False
Frame3.Visible = True
If optApply.Value = vbChecked Then Call optApply_Click

err:
If err <> 0 Then MsgBox err.Description
Me.MousePointer = vbDefault
ProgressBar1.Visible = False
Frame2.Visible = False
Close
End Sub

Private Sub CancelButton_Click()
frmOptions.Check3.Value = vbUnchecked
fraTitles.Visible = False
End Sub

Private Sub cmdLoad_Click()

If Dir("C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI") <> "" Then
    Dim logmode As Integer
    logmode = INIRead("LogSetting", "LogMode", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
    
    Select Case logmode
    Case 0
     '   Call ViewLogsOn(DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year)
        Call ViewLogsOn(DTPicker1.Value)
    Case 1
        txtLogname.Text = INIRead("LogSetting", "LogDir", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
        txtLogname.Text = txtLogname.Text & "\browse" & Format$(DTPicker1.Value, "ddmmyy") & "z" & INIRead("LogSetting", "extension", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
        Call OpenLog(txtLogname.Text)
    Case 2
        txtLogname.Text = INIRead("LogSetting", "LogDir", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
        txtLogname.Text = txtLogname.Text & "\browse" & INIRead("LogSetting", "extension", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
        Call OpenLog(txtLogname.Text)
    End Select
Else
    MsgBox "niL's KeyLogger have not been configured !", vbCritical
    cmdLoad.Enabled = False
    lblLoadAll.Enabled = False
End If
End Sub
Private Sub LoadUsername()
'get username

Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call getusername(sBuffer, lSize)
If lSize > 0 Then
        txtUsername = Left$(sBuffer, lSize)
Else
        txtUsername = vbNullString
End If

End Sub
Private Sub ViewLogsOn(ThisDate As String)

txtLogname.Text = INIRead("LogSetting", "LogDir", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
File1.Path = txtLogname.Text
txtLogname.Text = ""

List1.Clear
File1.ListIndex = 0
List2.Clear

Label13.Caption = "Logs on : " & ThisDate
Dim i As Integer
For i = 0 To File1.ListCount - 1
    File1.ListIndex = i
    
Dim ltime, ldate, encCode, noneed As String
If Dir(File1.Path & "\" & File1.Filename) <> "" Then
    On Error Resume Next
    Open File1.Path & "\" & File1.Filename For Input As #1
    Input #1, noneed, ltime, encCode, ldate
    Close #1
Else
    MsgBox "Log Not Found", vbInformation: Exit Sub
End If

   If noneed = "›" Then 'confirms that it is log
                
        If ldate = ThisDate Then 'Match date with GivenDate
            List2.AddItem File1.Path & "\" & File1.Filename
            List1.AddItem ltime & " - " & Right(FileDateTime(File1.Path & "\" & File1.Filename), 11)
        End If
   End If
Next i
frmChooseLog.Visible = True


End Sub



Private Sub Form_Load()

DTPicker1.Value = Format$(Now, "m/d/yyyy")
Call LoadUsername
End Sub

Private Sub lblLoadAll_Click()

If Dir("C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI") = "" Then Exit Sub

txtLogname.Text = INIRead("LogSetting", "LogDir", "C:\Documents and Settings\" & txtUsername & "\Application Data\System\SPYXX.INI")
txtLogname.Text = txtLogname.Text
On Error Resume Next
File1.Path = txtLogname.Text
txtLogname.Text = ""

List1.Clear
File1.ListIndex = 0
List2.Clear


Label13.Caption = "All Logs : "
Dim i As Integer
For i = 0 To File1.ListCount - 1
    File1.ListIndex = i
    
Dim ltime, ldate, encCode, noneed As String
If Dir(File1.Path & "\" & File1.Filename) <> "" Then
    On Error Resume Next
    Open File1.Path & "\" & File1.Filename For Input As #1
    Input #1, noneed, ltime, encCode, ldate
    Close #1
Else
    MsgBox "Log Not Found", vbInformation: Exit Sub
End If

   If noneed = "›" Then

            List2.AddItem File1.Path & "\" & File1.Filename
            List1.AddItem ltime & " - " & Right(FileDateTime(File1.Path & "\" & File1.Filename), 11) & " - " & ldate

   End If
Next i
frmChooseLog.Visible = True

End Sub
Private Function decrypt(iput As String, code As Integer)
Dim i As Integer
For i = 1 To Len(iput)  '
decrypt = decrypt & Chr(Asc(Mid(iput, i, 1)) - code)
Next i
End Function

Private Sub cmdNextSession_Click()
txtFind = "-_-_-_-_NEW SESSION-_-_-_-_-"
Call cmdFindNext_Click
End Sub

Private Sub cmdOK_Click()
If unipath <> "" Then
    Call OpenLog(unipath)
Else
    Frame2.Visible = False
End If
End Sub

Private Sub cmdOpen_Click()
'dlg.FileName = "*.txt"
dlg.ShowOpen
 Call OpenLog(dlg.Filename)
End Sub


Private Sub cmdTime_Click()
If Len(txtMM) = 1 Then txtMM = "0" + txtMM
If txtHH & ":" & txtMM <> "" Then
    txtFile.SelStart = txtFile.SelStart + 2
    If InStr(txtFile.SelStart, txtFile, txtHH & ":" & txtMM) <> 0 Then
      txtFile.SelStart = InStr(txtFile.SelStart, txtFile, txtHH & ":" & txtMM) - 1
      txtFile.SelLength = Len(txtHH & ":" & txtMM)
    Else
      MsgBox "Not found : " & txtHH & ":" & txtMM & vbCrLf
    End If
  End If
End Sub

Private Sub cmdUpdate_Click()
frmOptions.Show
End Sub



Private Sub Command2_Click()
Frame2.Visible = True
upwd.SetFocus
End Sub


Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
On Error GoTo err
txtFile.Width = Me.Width - 3600
txtFile.Height = Me.Height - 2000
Label10.Width = txtFile.Width
Call setFrames
err:
If err.Number = 380 Then MsgBox "Size too small!": Me.Width = (Screen.Width / 2): Me.Height = (Screen.Height / 2)
End Sub
Private Sub setFrames()
'Centre the Frames
frmChooseLog.Left = (Me.Width / 2) - (frmChooseLog.Width / 2)
Frame2.Left = (Me.Width / 2) - (Frame2.Width / 2)
frmChooseLog.Top = (Me.Height / 2) - (frmChooseLog.Height / 2)
Frame2.Top = (Me.Height / 2) - (Frame2.Height / 2)

Frame3.Top = txtFile.Top
Frame3.Left = txtFile.Width + txtFile.Left - Frame3.Width - 300
fraTitles.Top = Frame3.Top + Frame3.Height
fraTitles.Left = txtFile.Width + txtFile.Left - fraTitles.Width - 300
End Sub

Private Sub cmdFindNext_Click()
txtFind.AddItem txtFind.Text
  If txtFind <> "" Then
    txtFile.SelStart = txtFile.SelStart + 2
    If InStr(txtFile.SelStart, txtFile, txtFind) <> 0 Then
      txtFile.SelStart = InStr(txtFile.SelStart, txtFile, txtFind) - 1
      txtFile.SelLength = Len(txtFind)
    Else
      MsgBox "Not found : " & txtFind & vbCrLf
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

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Label12_Click()
frmChooseLog.Visible = False
End Sub



Private Sub Label9_Click()
unipath = ""
Frame2.Visible = False
End Sub

Private Sub List1_DblClick()

List2.ListIndex = List1.ListIndex

Call OpenLog(List2.Text)
End Sub

Private Sub optApply_Click()
If optApply.Value = vbChecked Then
    If Check1.Value = vbChecked Then Call ReplaceText("[C]", vbCrLf)
    If Check2.Value = vbChecked Then Call ReplaceText("[ENTR]", vbCrLf)
    If Check3.Value = vbChecked Then Call ReplaceText("[TAB]", vbTab)
    optApply.Caption = "&Reset"
Else
optApply.Caption = "&Apply"
Call OpenLog(txtLogname)
End If
End Sub

Private Sub Timer1_Timer()
If Label7.ForeColor = &HFF00& Then
Label7.ForeColor = &H8000&
Else: Label7.ForeColor = &HFF00&
End If

End Sub

Private Sub txtHH_Change()
If IsNumeric(txtHH) = False Or Val(txtHH) > 12 Then txtHH = "12"
End Sub

Private Sub txtMM_Change()
If IsNumeric(txtMM) = False Or Val(txtMM) > 59 Then txtMM = "00"
End Sub
Public Sub LoadTitles()

Dim t1, t2, t3, t4, t5 As String
On Error GoTo err

If Dir("C:\Documents and Settings\" & txtUsername.Text & "\Application Data\System\default.MCP") <> "" Then
    Open "C:\Documents and Settings\" & txtUsername.Text & "\Application Data\System\default.MCP" For Input As 1
    On Error Resume Next
    Input #1, t1, t2, t3, t4, t5
    Close #1

    txtTitle1 = t1
    txtTitle2 = t2
    txtTitle3 = t3
    txtTitle4 = t4
    txtTitle5 = t5
Else
    Call CreateTitles
    Call LoadTitles
End If

err:
Close   'Close all opened files
If err <> 0 Then MsgBox err.Description

End Sub
Private Sub CreateTitles()
Close
Open "C:\Documents and Settings\" & txtUsername.Text & "\Application Data\System\default.MCP" For Output As 1
Write #1, "Title1", "Title2", "Title3", "Title4", "Title5"
Close #1
End Sub

Private Sub OKButton_Click() ' Saves titles

    Open "C:\Documents and Settings\" & txtUsername.Text & "\Application Data\System\default.MCP" For Output As 1
    Write #1, txtTitle1, txtTitle2, txtTitle3, txtTitle4, txtTitle5
    Close #1
    
    fraTitles.Visible = False
    frmOptions.Show


End Sub



