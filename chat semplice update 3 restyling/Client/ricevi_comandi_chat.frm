VERSION 5.00
Begin VB.Form ricevi_comandi_chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "ricevi comandi chat"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdnascondi 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
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
      Caption         =   "nascondi"
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
   Begin client.CandyButton Cmdchiudi 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
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
   Begin VB.ListBox List1 
      BackColor       =   &H8000000A&
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "ricevi_comandi_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdchiudi_Click()
 Unload ricevi_comandi_chat
End Sub
