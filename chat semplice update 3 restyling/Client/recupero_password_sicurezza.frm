VERSION 5.00
Begin VB.Form recupero_password_sicurezza 
   BackColor       =   &H80000013&
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdsalva 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2280
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
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "quale' il tuo anno di nascita?"
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
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "se smarrisci la password prova a recuperarla rispondendo alla domanda.....ricordati la risposta"
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
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "recupero_password_sicurezza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdsalva_Click()
 Call SaveControlValues
End Sub

Private Sub form_load()
 Text1 = RegLoad(Text1)
End Sub

Private Sub SaveControlValues()
 Call RegSave(Text1, Text1.Text)
End Sub


