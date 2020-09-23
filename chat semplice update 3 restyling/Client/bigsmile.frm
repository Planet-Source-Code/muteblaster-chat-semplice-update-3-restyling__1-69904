VERSION 5.00
Begin VB.Form bigsmile 
   BorderStyle     =   0  'None
   Caption         =   "bigsmile"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Framebigsmile 
      BackColor       =   &H80000009&
      Caption         =   "big smile"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Image Image1 
         Height          =   1215
         Index           =   17
         Left            =   4200
         Picture         =   "bigsmile.frx":0000
         Tag             =   ":bigyessir"
         Top             =   4800
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   16
         Left            =   2280
         Picture         =   "bigsmile.frx":0B49
         Tag             =   ":bigyes"
         Top             =   4560
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1185
         Index           =   15
         Left            =   480
         Picture         =   "bigsmile.frx":1811
         Top             =   4800
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1185
         Index           =   14
         Left            =   7320
         Picture         =   "bigsmile.frx":22CB
         Tag             =   ":bigwoo"
         Top             =   3360
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1170
         Index           =   12
         Left            =   4080
         Picture         =   "bigsmile.frx":2D85
         Tag             =   ":bigsad"
         Top             =   3360
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   1485
         Index           =   4
         Left            =   7200
         Picture         =   "bigsmile.frx":39B6
         Tag             =   ":bigeat"
         Top             =   240
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   8
         Left            =   5640
         Picture         =   "bigsmile.frx":471D
         Tag             =   ":bighot"
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1350
         Index           =   18
         Left            =   5760
         Picture         =   "bigsmile.frx":5368
         Tag             =   ":bigzzz"
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   13
         Left            =   5520
         Picture         =   "bigsmile.frx":5DFA
         Tag             =   ":bigvomit"
         Top             =   3120
         Width           =   1725
      End
      Begin VB.Image Image1 
         Height          =   1590
         Index           =   3
         Left            =   5640
         Picture         =   "bigsmile.frx":67E8
         Tag             =   ":bigclassic"
         Top             =   120
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1575
         Index           =   9
         Left            =   7320
         Picture         =   "bigsmile.frx":714F
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   11
         Left            =   2160
         Picture         =   "bigsmile.frx":7C1F
         Tag             =   ":bigno"
         Top             =   3120
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   7
         Left            =   3960
         Picture         =   "bigsmile.frx":884D
         Tag             =   ":bigheart"
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1320
         Index           =   10
         Left            =   480
         Picture         =   "bigsmile.frx":90D3
         Top             =   3360
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   1110
         Index           =   6
         Left            =   2280
         Picture         =   "bigsmile.frx":9AEB
         Tag             =   ":bigyes"
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   5
         Left            =   360
         Picture         =   "bigsmile.frx":A630
         Tag             =   ":bigeye"
         Top             =   1680
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   1350
         Index           =   2
         Left            =   3840
         Picture         =   "bigsmile.frx":B142
         Tag             =   ":bigbored"
         Top             =   240
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1200
         Index           =   1
         Left            =   1920
         Picture         =   "bigsmile.frx":BB59
         Tag             =   ":bigasd"
         Top             =   240
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1650
         Index           =   0
         Left            =   120
         Picture         =   "bigsmile.frx":C75E
         Top             =   240
         Width           =   1650
      End
   End
End
Attribute VB_Name = "bigsmile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim frmResize As New ControlResizer

Private Sub form_load()
  
  frmResize.KeepRatio = True
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
    
End Sub
Private Sub Form_Resize()

  Call frmResize.FormResized(Me)
    
End Sub

Public Function FindAndReplace()
' Funzione FindAndReplace - By Xaxak

Dim FPos As Long
Dim FLen As Long

' :bigyes ---
FPos = chat.txtchat.Find(":bigyes", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 7
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigyes.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 8, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' :bigapplause ---
FPos = chat.txtchat.Find(":bigapplause", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 12
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigapplause.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 13, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigasd ---
FPos = chat.txtchat.Find(":bigasd", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 7
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigasd.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 8, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigbored ---
FPos = chat.txtchat.Find(":bigbored", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 9
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigbored.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 10, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigzzz ---
FPos = chat.txtchat.Find(":bigzzz", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 7
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigzzz.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 8, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigyessir ---
FPos = chat.txtchat.Find(":bigyessir", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 9
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigyessir.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 10, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigwoo ---
FPos = chat.txtchat.Find(":bigwoo", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 7
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigwoo.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 8, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigvomit ---
FPos = chat.txtchat.Find(":bigvomit", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 9
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigvomit.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 10, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigsad ---
FPos = chat.txtchat.Find(":bigsad", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 7
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigsad.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 8, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigno ---
FPos = chat.txtchat.Find(":bigno", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 6
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigno.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 7, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bighot ---
FPos = chat.txtchat.Find(":bighot", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 6
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bighot.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 7, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigheart ---
FPos = chat.txtchat.Find(":bigheart", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 9
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigheart.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 10, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bighaha ---
FPos = chat.txtchat.Find(":bigyes", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 8
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bighaha.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 9, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigeye ---
FPos = chat.txtchat.Find(":bigeye", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 6
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigeye.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 7, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigeat ---
FPos = chat.txtchat.Find(":bigeat", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 6
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigeat.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 7, , rtfNoHighlight)
chat.txtchat.Refresh
Wend

' bigclassic ---
FPos = chat.txtchat.Find(":bigclassic", 0, , rtfNoHighlight)
While FPos > 0
chat.txtchat.SelStart = FPos
chat.txtchat.SelLength = 11
chat.txtchat.SelText = ""
chat.txtchat.OLEObjects.Add , , App.Path & "\bigsmile\bigclassic.bmp"
DoEvents
FPos = chat.txtchat.Find(FString, FPos + 12, , rtfNoHighlight)
chat.txtchat.Refresh
Wend
End Function

Private Sub Image1_Click(Index As Integer)
 chat.Txtsend.Text = chat.Txtsend.Text & Image1(Index).Tag
 chat.Picture4.Top = 11400
End Sub
