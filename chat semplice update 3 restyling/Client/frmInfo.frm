VERSION 5.00
Begin VB.Form frmInfo 
   ClientHeight    =   3090
   ClientLeft      =   4935
   ClientTop       =   3690
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    txtInfo.Left = 0
    txtInfo.Top = 0
    txtInfo.Width = Me.ScaleWidth
    txtInfo.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtInfo = ""
    Me.Caption = ""
End Sub
