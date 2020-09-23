VERSION 5.00
Begin VB.Form animazione_connessione 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_animazione_connessione 
      Enabled         =   0   'False
      Interval        =   32
      Left            =   360
      Top             =   600
   End
   Begin VB.PictureBox Picture3 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   0
      Width           =   855
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         Picture         =   "animazione_connessione.frx":0000
         ScaleHeight     =   540
         ScaleWidth      =   23040
         TabIndex        =   2
         Top             =   0
         Width           =   23040
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         Picture         =   "animazione_connessione.frx":E4AE
         ScaleHeight     =   540
         ScaleWidth      =   23040
         TabIndex        =   1
         Top             =   0
         Width           =   23040
      End
   End
End
Attribute VB_Name = "animazione_connessione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer_animazione_connessione_Timer()
  Picture1.Left = Picture1.Left - 720
  Picture2.Left = Picture2.Left - 720
 If Picture1.Left <= -23040 Then
  Picture1.Left = Picture2.Left + Picture2.Width
 End If
If Picture2.Left <= -23040 Then
 Picture2.Left = Picture1.Left + Picture1.Width
End If
End Sub
