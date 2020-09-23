VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "Add Buddy"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtBuddy 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    AddBuddy (txtBuddy)
    txtBuddy.Text = ""
End Sub
