VERSION 5.00
Begin VB.Form popup 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   6150
   ClientTop       =   4815
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   840
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Labelallert 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Labeltitle 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   0
      Picture         =   "popup.frx":0000
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' popup di allerta piu' semplice di cosi' non si puo'
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim PosX As Long
Dim PosY As Long

Private Sub form_load()
MakeTransparent Me.Hwnd, 220
PosX = Screen.Width - Me.Width
PosY = Screen.Height
Me.Left = PosX
Me.Top = PosY
End Sub

' IL PRIMO TIMER REGOLA LA PARTENZA'
Private Sub Timer1_Timer()
Me.Top = Me.Top - 30
If Me.Top < PosY - Me.Height Then
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub
' IL SECONDO REGOLA LA DURATA'
Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer3.Enabled = True
End Sub
' IL TERZO CHIUDE IL POPUP'
Private Sub Timer3_Timer()
Me.Top = Me.Top + 30
If Me.Top = PosY Then
Timer3.Enabled = False
Unload Me
End If
End Sub

Private Sub Txtallert_Change()

End Sub
