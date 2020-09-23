VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9d.ocx"
Begin VB.Form animazioni_flash_chat 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin client.CandyButton Cmdflash1 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "OK"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
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
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin VB.TextBox Txtflash 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin client.CandyButton Cmdsucessico 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2520
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
      Caption         =   ">>>>>"
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
   Begin client.CandyButton Cmdpreceente 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2520
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
      Caption         =   "<<<<<<"
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash1 
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   3375
      _cx             =   5953
      _cy             =   2355
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   480
      Top             =   480
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "animazioni_flash_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flashnumero As Integer

Private Sub Cmdchiudi_Click()
 Unload animazioni_flash_chat
End Sub

Private Sub form_load()
 Cmdsucessico_Click
End Sub

Private Sub Cmdpreceente_Click()
 If flashnumero > 1 Then flashnumero = flashnumero - 1
  PlayFlashMovie App.Path & "\animazioni flash" & "\chat" & "\flash" & flashnumero & ".swf"
  Txtflash.Text = "flash" & flashnumero
End Sub

Private Sub Cmdsucessico_Click()
 If flashnumero < 15 Then flashnumero = flashnumero + 1
  PlayFlashMovie App.Path & "\animazioni flash" & "\chat" & "\flash" & flashnumero & ".swf"
  Txtflash.Text = "flash" & flashnumero
End Sub

Private Sub Txtflash_Change()
 If Txtflash.Text = "flash1" Then
    Label1.Caption = " x - Files"
 ElseIf Txtflash.Text = "flash2" Then
    Label1.Caption = "anitre migratorie"
 End If
End Sub
