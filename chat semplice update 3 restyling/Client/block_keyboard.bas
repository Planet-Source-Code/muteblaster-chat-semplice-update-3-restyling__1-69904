Attribute VB_Name = "block_keyboard"
Option Explicit

' ----------------------------------------------
'               Author   : Xip3000
'               eMail    : SuperXip@hotmail.com
'               Copyright: Xip3000
'-----------------------------------------------
' This Code Not disable CTRL-ALT-DELETE. Kill the Process when is enable.
'
' The code is for Windows XP.
'
'
' (*1) If you want it can eliminate this line. But the Task Manager didn't close.
'            SendKeys "%{F4}", True


Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                               ByVal nIDEvent As Long, _
                                               ByVal uElapse As Long, _
                                               ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long) As Long
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long

Private Window_to_Search As Long
Public tim As Integer

'-----------------------------------------------
'NAME OF WINDOW CAPTION TASK MANAGER
Public Task_Manager_Caption As String
'-----------------------------------------------



'-----------------------------------------------
  ' This Sub Search the Task Manager Process is show and Close it.
Public Sub TimerProc(ByVal lhwnd As Long, _
       ByVal uMsg As Long, _
       ByVal idEvent As Long, _
       ByVal dwTime As Long)

'-----------------------------------------------
  ' Search for the window caption is the TaskManager
    Window_to_Search = FindWindow(vbNullString, Task_Manager_Caption)
'-----------------------------------------------



    If Window_to_Search > 0 Then
'-----------------------------------------------
  ' Close the TaskManager. SendKey ALT-F4 to Task Manager Window
         SendKeys "%{F4}", True '(*1)*-------------->>> Read the in Code Head
'-----------------------------------------------
    
'-----------------------------------------------
  ' Set Form as Focus And new BlockInput (The TaskManager Deactivate the function)
         chat.SetFocus
         BlockInput True
'-----------------------------------------------
    End If
    
End Sub
