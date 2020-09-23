VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form copia_per_aiuto 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtnick 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CheckBox Checkremember 
      BackColor       =   &H80000013&
      Caption         =   "ricorda nick , password e frase tipica"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox Checkautoconnect 
      BackColor       =   &H80000013&
      Caption         =   "connetti automaticamente all' avvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   4920
      Width           =   255
   End
   Begin VB.CheckBox Checkattivasfondi_txtchat 
      BackColor       =   &H80000013&
      Caption         =   "attiva sfondi chat"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   5280
      Width           =   255
   End
   Begin VB.TextBox Txtip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame framelogin 
      BackColor       =   &H80000013&
      Caption         =   "login"
      Height          =   135
      Left            =   4200
      TabIndex        =   0
      Top             =   3480
      Width           =   375
      Begin MSComctlLib.ImageList status 
         Left            =   240
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":0352
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":0674
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":0EEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":1C8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":2B52
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":3A18
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList anim_browser 
         Left            =   840
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":484A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":4E74
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":549E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":5AC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":60F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":671C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":6D46
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":7370
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":799A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":7FC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":85EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":8C18
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":9242
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":986C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":9E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":A4C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "copia_per_aiuto.frx":AAEA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin client.CandyButton Cmddisconnetti 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "disconnetti"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin client.CandyButton CandyButton_avatar 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "scegli avatar"
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
   Begin RichTextLib.RichTextBox Txtfrase 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   4080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483629
      TextRTF         =   $"copia_per_aiuto.frx":B114
   End
   Begin client.CandyButton Cmdannulla 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
      Caption         =   "annulla"
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
   Begin client.CandyButton cmdLogin 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "accedi"
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
   Begin client.CandyButton Cmdritorna 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "ritorna in chat"
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
   Begin client.Anim Anim1 
      Height          =   1215
      Left            =   1560
      TabIndex        =   13
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
   End
   Begin client.Anim Anim4 
      Height          =   735
      Left            =   3480
      TabIndex        =   14
      Top             =   8520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
   End
   Begin VB.Image Picavatar 
      Height          =   1455
      Left            =   1680
      Picture         =   "copia_per_aiuto.frx":B196
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "copia_per_aiuto.frx":BEB4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   840
      TabIndex        =   32
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label_tempo_di_connessione 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Cmdminimizza 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Cmdexit 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Labelnick 
      BackStyle       =   0  'Transparent
      Caption         =   "login"
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   1080
      TabIndex        =   27
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Labelfrase 
      BackStyle       =   0  'Transparent
      Caption         =   "frase tipica"
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "ricorda nick , password e frase tipica"
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "connetti automaticamente all' avvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "attiva sfondi chat"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Labelmioip 
      BackStyle       =   0  'Transparent
      Caption         =   "mio ip"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Labelip 
      BackStyle       =   0  'Transparent
      Caption         =   "indirizzo del server"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Cmdprofilo 
      BackStyle       =   0  'Transparent
      Caption         =   "crea profilo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label cmdCreate 
      BackStyle       =   0  'Transparent
      Caption         =   "crea account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label_vediaccount 
      BackStyle       =   0  'Transparent
      Caption         =   "vedi account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Cmdrouter 
      BackStyle       =   0  'Transparent
      Caption         =   "router"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Cmdtest 
      BackStyle       =   0  'Transparent
      Caption         =   "testa porte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   8160
      Width           =   1455
   End
End
Attribute VB_Name = "copia_per_aiuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
