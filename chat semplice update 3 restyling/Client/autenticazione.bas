Attribute VB_Name = "autenticazione"
Public LocalInfo As String
Public timesgotinfo As Integer

Private Type FLASHWINFO
  cbSize As Long
  Hwnd As Long
  dwFlags As Long
  uCount As Long
  dwTimeout As Long
End Type

Private Const FLASHW_TRAY = 2


Private Declare Function LoadLibrary Lib "kernel32" _
  Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, ByVal lpProcName As String) As Long


Private Declare Function FreeLibrary Lib "kernel32" _
  (ByVal hLibModule As Long) As Long

Private Declare Function FlashWindowEx Lib "user32" _
   (FWInfo As FLASHWINFO) As Boolean

Public sent As Boolean
Public imwindow(0 To 50) As Form

Public Sub FlashWindow(Hwnd As Long, _
  Optional NumberOfFlashes As Integer = 5)
'***************************************************
'Purpose: Flashes a Window in the taskbar in order to notify
'a user of an event within a program

'Parameters: Hwnd=hwnd of frm to flash
             'NumberofFlashes = Number of times to
               'flash

'Notes: WINDOWS 98 OR 2000 is REQUIRED

'Uses FlashWindowEx API, which  substitutes
'for bringing you window to the foreground
'obtrusively (e.g., on startup or when siginficant
'event occurs in your program) Windows 98/2000 no
'longer permits this

'Example:

'FlashWindow me.hwnd

'***************************************************
'Prevent Errors by checking if
'the API function is available on the
'Current OS

If Not APIFunctionPresent("FlashWindowEx", "user32") _
   Then Exit Sub

Dim bRet As Boolean
Dim udtFWInfo As FLASHWINFO

With udtFWInfo
   .cbSize = 20
   .Hwnd = Hwnd
   .dwFlags = FLASHW_TRAY
   .uCount = NumberOfFlashes 'flash window 5 times
   .dwTimeout = 0
End With

bRet = FlashWindowEx(udtFWInfo)
End Sub

Private Function APIFunctionPresent(ByVal FunctionName _
   As String, ByVal DllName As String) As Boolean

'USAGE:
'Dim bAvail as boolean
'bAvail = APIFunctionPresent("GetDiskFreeSpaceExA", "kernel32")

    Dim lHandle As Long
    Dim lAddr  As Long

    lHandle = LoadLibrary(DllName)
    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        FreeLibrary lHandle
    End If
    
    APIFunctionPresent = (lAddr <> 0)

End Function

Public Sub AddBuddy(buddytoadd As String)
    AddUser = True
    
    For I = 0 To frmBuddyList.lstBuddy.ListCount - 1
        If UCase(buddytoadd) = UCase(frmBuddyList.lstBuddy.List(I)) Then
            MsgBox "Buddy already on list!", , "Error!"
            AddUser = False
        End If
    Next
    For I = 0 To frmBuddyList.lstOffline.ListCount - 1
        If UCase(buddytoadd) = UCase(frmBuddyList.lstOffline.List(I)) Then
            MsgBox "Buddy already on list!", , "Error!"
            AddUser = False
        End If
    Next
    If buddytoadd <> "" And AddUser = True Then
        login.win.SendData "add-" & buddytoadd & "\-"
        
    End If
End Sub

