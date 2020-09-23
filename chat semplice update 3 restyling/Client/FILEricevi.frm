VERSION 5.00
Begin VB.Form FILEricevi 
   Caption         =   "ricevi file"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Server"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.ListBox Listavviso 
         Height          =   450
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Cmdchiudi 
         Caption         =   "Chiudi"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "FILEricevi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BlnTflag As Boolean ' Transfert Flag'
Dim LngCursor As Long ' source file position pointer'

Private Sub Cmdchiudi_Click()
login.Wsricevifile.Close
End Sub

