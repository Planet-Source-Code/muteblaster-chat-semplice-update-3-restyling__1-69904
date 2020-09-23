VERSION 5.00
Begin VB.Form smyle 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "smyle"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraIcons 
      BackColor       =   &H00FFFFFF&
      Caption         =   "smyle"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   5040
         TabIndex        =   1
         Top             =   840
         Width           =   135
      End
      Begin VB.Image imgIcon 
         Height          =   270
         Index           =   52
         Left            =   480
         Picture         =   "smyle.frx":0000
         Tag             =   "(li)"
         Top             =   720
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   270
         Index           =   51
         Left            =   120
         Picture         =   "smyle.frx":03D3
         Tag             =   "(st)"
         Top             =   720
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   50
         Left            =   1560
         Picture         =   "smyle.frx":0604
         Tag             =   "(R)"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   49
         Left            =   840
         Picture         =   "smyle.frx":0A05
         Tag             =   "(#)"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   48
         Left            =   1200
         Picture         =   "smyle.frx":0DF1
         Tag             =   "(*)"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   47
         Left            =   120
         Picture         =   "smyle.frx":11FD
         Tag             =   "(0)"
         Top             =   1080
         Width           =   390
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   46
         Left            =   2280
         Picture         =   "smyle.frx":1666
         Tag             =   "(t)"
         Top             =   1080
         Width           =   390
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   45
         Left            =   480
         Picture         =   "smyle.frx":1A5D
         Tag             =   "(ap)"
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   44
         Left            =   3000
         Picture         =   "smyle.frx":1E4C
         Tag             =   "(ip)"
         Top             =   720
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   43
         Left            =   2280
         Picture         =   "smyle.frx":225F
         Tag             =   "(um)"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   42
         Left            =   1920
         Picture         =   "smyle.frx":2660
         Tag             =   "(mp)"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   41
         Left            =   2640
         Picture         =   "smyle.frx":2A66
         Tag             =   "(W)"
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   40
         Left            =   2640
         Picture         =   "smyle.frx":2E30
         Tag             =   "(P)"
         Top             =   720
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   315
         Index           =   39
         Left            =   1560
         Picture         =   "smyle.frx":3271
         Tag             =   "(Z)"
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   315
         Index           =   38
         Left            =   1200
         Picture         =   "smyle.frx":3630
         Tag             =   "(X)"
         Top             =   1080
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   435
         Index           =   37
         Left            =   3360
         Picture         =   "smyle.frx":39F4
         Tag             =   "<:o)"
         Top             =   240
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   36
         Left            =   840
         Picture         =   "smyle.frx":3B25
         Tag             =   "(pi) <pizza>"
         Top             =   1080
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   270
         Index           =   35
         Left            =   4080
         Picture         =   "smyle.frx":3C62
         Tag             =   "(h)"
         Top             =   360
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   34
         Left            =   4440
         Picture         =   "smyle.frx":3F1F
         Tag             =   "(&)"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   33
         Left            =   5520
         Picture         =   "smyle.frx":4368
         Tag             =   "(8)"
         Top             =   720
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   32
         Left            =   5160
         Picture         =   "smyle.frx":47D0
         Tag             =   "(~)"
         Top             =   720
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   31
         Left            =   3360
         Picture         =   "smyle.frx":4924
         Tag             =   "(^) <cake> <torta>"
         Top             =   720
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   30
         Left            =   4800
         Picture         =   "smyle.frx":4E75
         Tag             =   "(u)"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   29
         Left            =   5520
         Picture         =   "smyle.frx":52F9
         Tag             =   "(bah)"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   28
         Left            =   5160
         Picture         =   "smyle.frx":5739
         Tag             =   "(so)"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   27
         Left            =   4800
         Picture         =   "smyle.frx":5A20
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   26
         Left            =   3720
         Picture         =   "smyle.frx":5B79
         Tag             =   "(g) <regalo>"
         Top             =   720
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   315
         Index           =   25
         Left            =   1920
         Picture         =   "smyle.frx":5E5B
         Tag             =   "(m) <2people>"
         Top             =   1080
         Width           =   330
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   24
         Left            =   3000
         Picture         =   "smyle.frx":62EB
         Tag             =   "8-) <eek>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   675
         Index           =   23
         Left            =   4800
         Picture         =   "smyle.frx":668E
         Tag             =   "(co) <pc>"
         Top             =   1560
         Width           =   660
      End
      Begin VB.Image imgIcon 
         Height          =   255
         Index           =   22
         Left            =   4080
         Picture         =   "smyle.frx":6D5A
         Tag             =   "(au) <car>"
         Top             =   720
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   21
         Left            =   3720
         Picture         =   "smyle.frx":6F4E
         Tag             =   ":-*"
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   4
         Left            =   1560
         Picture         =   "smyle.frx":70B3
         Tag             =   ":( :-( <triste>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   3
         Left            =   1200
         Picture         =   "smyle.frx":75E5
         Tag             =   ";) ;-) <occhiolino>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   2
         Left            =   840
         Picture         =   "smyle.frx":7B17
         Tag             =   ":P :-P <lingua>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   1
         Left            =   480
         Picture         =   "smyle.frx":8049
         Tag             =   ":D :-D <risata> <lol>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   0
         Left            =   120
         Picture         =   "smyle.frx":857B
         Tag             =   ":) :-) <sorriso>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Index           =   5
         Left            =   1920
         Picture         =   "smyle.frx":8AAD
         Tag             =   "8o| >:( >:-( <eg>"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   6
         Left            =   2280
         Picture         =   "smyle.frx":8FEF
         Tag             =   ":'( :,-( :,( <pianto>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   8
         Left            =   4200
         Picture         =   "smyle.frx":9521
         Tag             =   "(k) <kiss> <bacio>"
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   9
         Left            =   2280
         Picture         =   "smyle.frx":988E
         Tag             =   ":-o :-O <impress>"
         ToolTipText     =   "impressionato"
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   285
         Index           =   10
         Left            =   4440
         Picture         =   "smyle.frx":9C4A
         Tag             =   "(l) <love> <amore> <cuore>"
         Top             =   720
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   11
         Left            =   1680
         Picture         =   "smyle.frx":A0A7
         Tag             =   "(c) <caffè> <coffee>"
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   17
         Left            =   3720
         Picture         =   "smyle.frx":A656
         Tag             =   "<peluche> <orsetto> <pupazzo>"
         Top             =   1080
         Width           =   555
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   12
         Left            =   3000
         Picture         =   "smyle.frx":AC8A
         Tag             =   "(f) @>- <rosa> <flower> <fiore>"
         Top             =   1080
         Width           =   555
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   13
         Left            =   1200
         Picture         =   "smyle.frx":B236
         Tag             =   "(B) <birra> <beer>"
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   14
         Left            =   2880
         Picture         =   "smyle.frx":B55E
         Tag             =   "(S) <luna> <moon>"
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   555
         Index           =   15
         Left            =   120
         Picture         =   "smyle.frx":B8C7
         Tag             =   "<cibo> <fastfood> <food>"
         Top             =   1440
         Width           =   540
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   16
         Left            =   4320
         Picture         =   "smyle.frx":BCB0
         Tag             =   "(i) <idea> <!> <lampadina>"
         Top             =   1080
         Width           =   555
      End
      Begin VB.Image imgIcon 
         Height          =   225
         Index           =   7
         Left            =   2640
         Picture         =   "smyle.frx":C1AF
         Tag             =   "<pazzo> <crazy>"
         Top             =   360
         Width           =   225
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   18
         Left            =   5040
         Picture         =   "smyle.frx":C536
         Tag             =   "(ci) <sigaro> <havana>"
         Top             =   1080
         Width           =   570
      End
      Begin VB.Image imgIcon 
         Height          =   450
         Index           =   19
         Left            =   3480
         Picture         =   "smyle.frx":C8BD
         Tag             =   "(mo) <soldi> <$$$> <money>"
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image imgIcon 
         Height          =   600
         Index           =   20
         Left            =   720
         Picture         =   "smyle.frx":CED7
         Tag             =   "<vino> <wine>"
         Top             =   1440
         Width           =   450
      End
   End
End
Attribute VB_Name = "smyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Const WM_PASTE = &H302
Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Private MovePicture1 As Boolean
 
Private Sub imgIcon_Click(Index As Integer)
'Clipboard.Clear                                     'cancella il contenuto degli appunti
'Clipboard.SetText " "                               'copia uno spazio vuoto
'SendMessage chat.txtsend.hWnd, WM_PASTE, 0, 0&   'incolla il contenuto degli appunti
'Clipboard.Clear                                     'cancella il contenuto degli appunti
'Clipboard.SetData imgIcon(Index).Picture
'chat.txtsend.SetFocus
'chat.txtsend.SelStart = Len(chat.txtsend.Text)
'SendMessage chat.txtsend.hWnd, WM_PASTE, 0, 0&   'incolla il contenuto degli appunti
'Clipboard.Clear                                     'cancella il contenuto degli appunti per completare l'operazione.
                'cancella il contenuto degli appunti per completare l'operazione.
chat.txtsend.Text = chat.txtsend.Text & Space(1) & Mid(imgIcon(Index).Tag, 1, InStr(1, imgIcon(Index).Tag, " ") - 1)
chat.Picture3.Top = 10680
End Sub


 

