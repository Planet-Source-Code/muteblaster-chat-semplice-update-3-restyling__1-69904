VERSION 5.00
Begin VB.Form seleziona_database_immagine 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdimmagini_personali 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "immagini personali"
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
End
Attribute VB_Name = "seleziona_database_immagine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdimmagini_personali_Click()
  database_immagini_personali.Show
End Sub
