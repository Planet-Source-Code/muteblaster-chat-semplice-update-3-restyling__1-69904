VERSION 5.00
Begin VB.Form agenda 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "AGENDA"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton CMDSALVA 
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6840
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
      Caption         =   "salva"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   5535
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox Text10 
         Height          =   1335
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3480
         Width           =   5655
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   24
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "SCRIVI OPINIONI A RIGUARDO DELL'AMICO"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CELLULARE"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "INDIRIZZO"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6735
      Begin client.CandyButton CandyButton1 
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">>>>>"
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
      Begin VB.TextBox Text5 
         Height          =   1335
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3480
         Width           =   5415
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   25
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SCRIVI COMMENTO A RIGUARDO DELL'AMICO"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CELLULARE"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "INDIRIZZO"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
   Begin client.CandyButton Cmdexit 
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
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
   Begin VB.Shape Shape1 
      Height          =   7335
      Left            =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   0
      Picture         =   "agenda.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   6600
      Picture         =   "agenda.frx":08B2
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   240
      Picture         =   "agenda.frx":0FC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6405
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "AGENDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   26
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private OldX As Integer
Private OldY As Integer

Private Sub form_load()
 Text1 = RegLoad(Text1)
 Text2 = RegLoad(Text2)
 Text3 = RegLoad(Text3)
 Text3 = RegLoad(Text3)
 Text4 = RegLoad(Text4)
 Text5 = RegLoad(Text5)
 Text6 = RegLoad(Text6)
 Text7 = RegLoad(Text7)
 Text8 = RegLoad(Text8)
 Text9 = RegLoad(Text9)
 Text10 = RegLoad(Text10)
End Sub

Private Sub SaveControlValues()
 Call RegSave(Text1, Text1.Text)
 Call RegSave(Text2, Text2.Text)
 Call RegSave(Text3, Text3.Text)
 Call RegSave(Text4, Text4.Text)
 Call RegSave(Text5, Text5.Text)
 Call RegSave(Text6, Text6.Text)
 Call RegSave(Text7, Text7.Text)
 Call RegSave(Text8, Text8.Text)
 Call RegSave(Text9, Text9.Text)
 Call RegSave(Text10, Text10.Text)
End Sub

Private Sub Cmdsalva_Click()
 Call SaveControlValues
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage agenda.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
