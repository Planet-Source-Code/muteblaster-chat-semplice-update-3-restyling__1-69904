Attribute VB_Name = "listboxmod01"
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194

Public Sub AddHScroll(List As ListBox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'trova il testo piu' lungo nella lista'
    For i = 0 To List.ListCount - 1
        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(1))
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'usiamo le api '
    SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub



