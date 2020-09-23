Attribute VB_Name = "animazioni_flash"
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
    Public Const SW_NORMAL = 1
        
Public Sub OpenWebsite(strWebsite As String)
    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
    End If
End Sub

Public Function PlayFlashMovie(Filename As String)
  With animazioni_flash_chat.flash1
      .Movie = Filename
      .Play
  End With
End Function
