VERSION 5.00
Begin VB.Form avatar 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "avatar"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdok 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ok"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Txtavatar 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Cmdprecedente 
      BackColor       =   &H80000013&
      Caption         =   "<<<<"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Cmdsucessivo 
      BackColor       =   &H80000013&
      Caption         =   ">>>>"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Picavatar 
      Height          =   1455
      Left            =   240
      Picture         =   "avatar.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "avatar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Dim immaginenumero As Integer ' dichiariamo la variabile che identifica il numero delle immagini'

Private Sub Cmdok_Click()
 login.Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & avatar.Txtavatar.Text & ".gif")
 frmBuddyList.Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & avatar.Txtavatar.Text & ".gif")
' login.Picture5.Top = 9480  '
Unload avatar
End Sub

'all'avvio mettiamo attivo il comando di avanzamento delle immagini'
' nessuna importanza solo che cosi' il form parte con una immagine pronta'
Private Sub Form_Load()
    Cmdsucessivo_Click
    Cmdprecedente_Click
End Sub

' creeiamo un bottone per poter passarae da una immagine alla sucessiva ho previsto'
' solo gif ma ovviamente si possono mettere anche jpg, bisogna specificarlo'
Private Sub Cmdprecedente_Click()
If immaginenumero > 1 Then immaginenumero = immaginenumero - 1 ' bisogna indicare il numero massimo di gif presenti'
                                                               ' altrimenti continuando a proseguire va' in crash'
Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & immaginenumero & ".gif") ' mettiamo nella picturebox l'immagine scelta'
Txtavatar.Text = immaginenumero
End Sub

' questo comando serve per tornare indietro'
Private Sub Cmdsucessivo_Click()
If immaginenumero < 79 Then immaginenumero = immaginenumero + 1
Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & immaginenumero & ".gif")
Txtavatar.Text = immaginenumero
End Sub


