Attribute VB_Name = "Module1"

Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As _
Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal _
nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As _
Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Sub TransparentForm(frm As Form)
    frm.ScaleMode = vbPixels
    Const RGN_DIFF = 4
    Const RGN_OR = 2

    Dim outer_rgn As Long
    Dim inner_rgn As Long
    Dim wid As Single
    Dim hgt As Single
    Dim border_width As Single
    Dim title_height As Single
    Dim ctl_left As Single
    Dim ctl_top As Single
    Dim ctl_right As Single
    Dim ctl_bottom As Single
    Dim control_rgn As Long
    Dim combined_rgn As Long
    Dim ctl As Control

    If frm.WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, _
        hgt - border_width)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF


    'Restrict the window to the region.
    SetWindowRgn frm.hwnd, combined_rgn, True
End Sub






