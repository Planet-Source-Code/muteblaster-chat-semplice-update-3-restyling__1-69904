VERSION 5.00
Begin VB.Form licenza 
   BackColor       =   &H80000013&
   Caption         =   "GNU"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   6000
      Picture         =   "licenza.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.ListBox ListGNU 
      BackColor       =   &H8000000A&
      Height          =   7080
      ItemData        =   "licenza.frx":14CD
      Left            =   360
      List            =   "licenza.frx":181C
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YOU ARE ALOWED TO MODIFY BUT YOU ARE NOT ALOWED TO SELL FOR PROFIT "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label LabelGNU 
      BackStyle       =   0  'Transparent
      Caption         =   "QUESTO SOFTWARE E' PROTETTO DALLA LICENZA GNU"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "licenza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
Call AddHScroll(ListGNU)
End Sub

