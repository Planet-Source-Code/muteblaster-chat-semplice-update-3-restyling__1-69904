VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form multiroom 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Chat Rooms"
      Height          =   3375
      Left            =   4560
      TabIndex        =   8
      Top             =   1200
      Width           =   1995
      Begin VB.ListBox List2 
         Height          =   2985
         ItemData        =   "multiroom.frx":0000
         Left            =   165
         List            =   "multiroom.frx":000D
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   105
      TabIndex        =   5
      Top             =   240
      Width           =   4275
      Begin VB.TextBox Txtchatroom 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1755
         TabIndex        =   7
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "chat room"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "nick"
         Height          =   240
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Height          =   450
      Left            =   5640
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   450
      Left            =   4680
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send"
      Height          =   450
      Left            =   3690
      TabIndex        =   2
      Top             =   4695
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   4740
      Width           =   3570
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   30
      TabIndex        =   0
      Top             =   1290
      Width           =   4380
   End
   Begin MSWinsockLib.Winsock client1 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "multiroom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim chatname As String
Private Sub Form_Load()
 Text3.Text = login.Txtnick.Text
  Command2_Click
End Sub

Private Sub client1_Connect()
client1.SendData "<COMMAND>" & Text3.Text
End Sub

Private Sub client1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
Dim operation As String
client1.GetData strdata
If Left(strdata, 9) = "<COMMAND>" Then
operation = Mid(strdata, 10, Len(strdata) - 9)
command operation
Exit Sub
End If
List1.AddItem strdata
List1.ListIndex = List1.NewIndex
End Sub


Private Sub Command1_Click()
On Error Resume Next
client1.SendData Text3.Text & " = " & Text1.Text
Text1.Text = ""
End Sub

Private Sub Command2_Click()
client1.RemoteHost = login.txtip.Text
client1.RemotePort = 1001
client1.Connect
End Sub

Private Sub Command3_Click()
On Error Resume Next
client1.SendData "<COMMAND>dis"
DoEvents
client1.Close
End Sub



Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
 multiroom.Visible = False
End Sub

Private Sub List2_Click()
client1.SendData "<COMMAND>crm" & List2.ListIndex + 1
List1.Clear
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
End If
End Sub


Private Sub Timer1_Timer()
Select Case client1.State
Case Is = 7
Me.Caption = "Connected to server"
Case Is = 9
Me.Caption = "Disconnected"
Command3_Click
Case Is = 0
Me.Caption = "Disconnected"
Command3_Click
End Select
End Sub

Public Sub command(commandop As String)
Select Case commandop
Case Is = "dis"
client1.Close
End Select
End Sub
