VERSION 5.00
Begin VB.Form profiloutente 
   BorderStyle     =   0  'None
   Caption         =   "mio profilo"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frameprofiloutente 
      BackColor       =   &H8000000A&
      Caption         =   "questo e' il mio profilo"
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin client.CandyButton Cmdchiudi 
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   5880
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "chiudi"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.TextBox Txtchatnick 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2880
         Width           =   2055
      End
      Begin VB.PictureBox Picavatar 
         BackColor       =   &H8000000D&
         Height          =   1455
         Left            =   480
         ScaleHeight     =   1395
         ScaleWidth      =   1515
         TabIndex        =   7
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Frame Frameconnessione 
         BackColor       =   &H8000000D&
         Caption         =   "connessione"
         Height          =   2295
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4095
         Begin VB.TextBox Txtfrasetipica 
            Appearance      =   0  'Flat
            Height          =   645
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Txtserver 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox Txtnick 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "frase tipica"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "nome del server"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "nick del login"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Shape Shape2 
         Height          =   1455
         Left            =   4560
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   1440
         Left            =   4560
         Picture         =   "profiloutente.frx":0000
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "nick in chat"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2880
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   1695
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "avatar attuale"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   4080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "profiloutente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdchiudi_Click()
 chat.Picture15.Top = 10800
End Sub

Private Sub form_load()
 Txtnick.Text = login.Txtnick.Text
 Txtserver.Text = login.txtIP.Text
 Txtchatnick.Text = chat.Txtmionick.Text
 Txtfrasetipica.Text = login.Txtfrase.Text
 Picavatar.Picture = LoadPicture(App.Path & "\avatar" & "\immagine" & avatar.Txtavatar & ".gif") ' richiamiamo l'immagine in base all'indice'
End Sub
