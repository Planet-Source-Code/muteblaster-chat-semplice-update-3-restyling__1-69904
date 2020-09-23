VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Public Title As String
Public Message As String
Public PosX As Long
Public PosY As Long

Public Sub MSN6()
Dim alert As New popup
alert.Labeltitle.Caption = Title
alert.Labelallert.Caption = Message
alert.Show
End Sub

