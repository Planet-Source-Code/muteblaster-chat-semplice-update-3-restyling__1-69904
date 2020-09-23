Attribute VB_Name = "trasparent_richtxtbox"
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TRANSPARENT = &H20&

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub SetTransparent(ByVal mHWnd As Long)
    Dim lS As Long
    If mHWnd <> 0 Then
       lS = GetWindowLong(mHWnd, GWL_EXSTYLE)
       lS = lS Or WS_EX_TRANSPARENT
       SetWindowLong mHWnd, GWL_EXSTYLE, lS
    End If
End Sub

