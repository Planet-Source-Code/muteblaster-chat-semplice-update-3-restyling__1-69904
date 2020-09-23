VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRooms 
   BorderStyle     =   0  'None
   Caption         =   "Chat Example - Rooms"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin client.CandyButton cmdCreate 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "crea canale"
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
   Begin VB.Frame fraRooms 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin MSComctlLib.TreeView lstRooms 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   4895
         _Version        =   393217
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
Dim Msg As String
    Msg = InputBox("Please type the name of the room you want to create:", "Create a room")
    If Msg = "" Then
        Exit Sub
    ElseIf Msg <> "" Then
        Call SendData(CreateRoom(Msg))
        lstRooms.Nodes.Add , , , Msg & " (0)"
    End If
End Sub
Private Sub lstRooms_DblClick()
Dim Rooms As String
    Rooms = Split(lstRooms.SelectedItem.Text, " (")(0)
    Call SendData(JoinRoom(Rooms, ChatUser))
    With frmChat
        .Caption = "Chat Room -- " & Rooms
        .txtChat.Text = ""
        .lstUsers.Nodes.Clear
        Call ChatEntry(Rooms, "Chat with new people, Enjoy!")
        .Show
    End With
    chat.Timer_chiusura_frmrooms.Enabled = True
End Sub
