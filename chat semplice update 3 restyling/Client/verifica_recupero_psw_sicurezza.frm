VERSION 5.00
Begin VB.Form verifica_recupero_psw_sicurezza 
   BackColor       =   &H80000013&
   Caption         =   "verifica"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdverifica 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "verifica"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "scrvi la risposta"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "quale' la tua data di nascita?"
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
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "la domanda era"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "verifica_recupero_psw_sicurezza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdverifica_Click()
If Text1.Text = recupero_password_sicurezza.Text1.Text Then
   login.Frame1.Visible = True
   login.framelogin.Visible = True
   login.Frame2.Top = 600
   login.Frame2.Left = 240
   login.Frame4.Visible = False
   Unload verifica_recupero_psw_sicurezza
Else
   avviso.Show
   avviso.Labelmessaggio.Caption = "la risposta e' sbagliata"
   Text1.Text = ""
 End If
End Sub
