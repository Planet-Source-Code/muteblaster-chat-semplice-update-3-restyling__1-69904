VERSION 5.00
Begin VB.Form frmSetInfo 
   Caption         =   "Change Info"
   ClientHeight    =   3105
   ClientLeft      =   4935
   ClientTop       =   3510
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4755
   Visible         =   0   'False
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Info"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
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
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmSetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    login.win.SendData "sinf-" & txtInfo.Text
    LocalInfo = txtInfo.Text & vbCrLf
End Sub
