Attribute VB_Name = "Declarations"
'******************************************************************************************
'     Sample for retrieving keystrokes  by use of the "kbLog32.dll"
'                      (c) 2002 by Nilesh Akhade.
'******************************************************************************************
Option Explicit
'******************************************************************************************
'DLL declarations
Public Declare Function StartLog Lib "kbLog32" (ByVal hWnd As Long, _
                            ByVal lpFuncAddress As Long) As Long

Public Declare Sub EndLog Lib "kbLog32" ()

'----------------------------------------------------------------------------------------
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'****************************************************************************************
' Keyboard messages
Public Const WM_KEYUP = &H101
Public Const WM_KEYDOWN = &H100
Public Const WM_CHAR = &H102
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105

'SetWindowPos messages
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
'******************************************************************************************


