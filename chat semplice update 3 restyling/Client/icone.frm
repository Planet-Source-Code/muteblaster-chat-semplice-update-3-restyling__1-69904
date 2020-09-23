VERSION 5.00
Begin VB.Form icone 
   Caption         =   "icone"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2520
      Top             =   1440
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   720
      Picture         =   "icone.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "icone.frx":058A
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   1
      Left            =   600
      Picture         =   "icone.frx":0B14
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   2
      Left            =   960
      Picture         =   "icone.frx":109E
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   3
      Left            =   1320
      Picture         =   "icone.frx":1628
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   4
      Left            =   1680
      Picture         =   "icone.frx":1BB2
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   5
      Left            =   2040
      Picture         =   "icone.frx":213C
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image_notifica 
      Height          =   240
      Index           =   6
      Left            =   2400
      Picture         =   "icone.frx":26C6
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "icone.frx":2C50
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "icone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
 Timer2.Enabled = True
 Static mIcon As Long
    If mIcon = 7 Then mIcon = 0
    TrayModify Tray_Icon, Image_notifica(mIcon).Picture
    mIcon = mIcon + 1
End Sub

Private Sub Timer2_Timer()
 Timer1.Enabled = False
 TrayModify Tray_Icon, Image1.Picture
 Timer2.Enabled = False
End Sub
