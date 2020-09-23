Attribute VB_Name = "muovi_componenti"
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function movey(Mbutton As Integer, FRM As Form, CRL As Control, FRMhasMenu As Boolean)
'Mbutton is to check if mouse is down, FRM to get the form, CRL to get the control to move, FRMhasMenu is for ajusting form top
If Mbutton = 1 Then
Dim cursorPos As POINTAPI
GetCursorPos cursorPos
If FRMhasMenu = True Then
CRL.Top = (cursorPos.Y * Screen.TwipsPerPixelY - FRM.Top) - 725
Else
CRL.Top = (cursorPos.Y * Screen.TwipsPerPixelY - FRM.Top) - 500
End If
End If
End Function
Public Function movex(Mbutton As Integer, FRM As Form, CRL As Control)
If Mbutton = 1 Then
Dim cursorPos As POINTAPI
GetCursorPos cursorPos
CRL.Left = (cursorPos.X * Screen.TwipsPerPixelX - FRM.Left) - 1400
End If
End Function
