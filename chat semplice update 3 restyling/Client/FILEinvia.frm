VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FILEinvia 
   Caption         =   "Disconnected"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcriptafile 
      Caption         =   "cripta file"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "connetti"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Cscegli 
      Caption         =   "sceglifile"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Framecontatti 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.TextBox TxtCPort 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "3335"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Txtmionick 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblPort 
         Caption         =   "Port"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton CmdSend 
      Cancel          =   -1  'True
      Caption         =   "Send File"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Cmddisconnect 
      Caption         =   "disconnetti"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog DlgSend 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsinviofile 
      Left            =   1920
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Labelcripta 
      BackStyle       =   0  'Transparent
      Caption         =   "cripta il file prima di spedirlo"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label LblStatut 
      Caption         =   "Statut:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   8415
   End
End
Attribute VB_Name = "FILEinvia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------'
'    QUESTA PARTE DI CODICE NON E' FATTA DA ME' MA DA UN MIO AMICO PROFESSIONISTA       '
'    si e' sforzato per fare il file transfer cosi' semplice :)                         '
'---------------------------------------------------------------------------------------'
Option Explicit
Dim BlnTflag As Boolean '' Transfert Flag'
Dim LngCursor As Long '' source file position pointer'

Private Sub Form_Load()
 Text2.Text = chat.Text2.Text
Wsinviofile.connect Trim(chat.txtIpUtente.Text), TxtCPort.Text
End Sub
Private Sub cmdConnect_Click()
On Error Resume Next
Wsinviofile.connect Trim(chat.txtIpUtente.Text), TxtCPort.Text 'Connection'
End Sub
Private Sub Cmddisconnect_Click()
Wsinviofile.Close
Unload FILEinvia
End Sub

Private Sub Cmdsend_Click()
On Error GoTo actcancel

If BlnTflag = False Then '' if we're not transferring'
    LngCursor = 0 ' pointer reinitialisation'
    If Wsinviofile.State <> 7 Then
        'Call ErrorHandler(2)'
    Else
        Wsinviofile.SendData "Transfert" & "|" & DlgSend.FileTitle & "|" & FileLen(DlgSend.FileName) 'on envoie le nom du fichier, we send the file name
    End If
Else
    Exit Sub
End If

actcancel: Exit Sub
End Sub

Private Sub Cscegli_Click()
DlgSend.ShowOpen 'common dialog'
End Sub

Private Sub Wsinviofile_Close()
avviso.Show
 avviso.Labelmessaggio.Caption = " la connessione e' chiusa"
Wsinviofile.Close
End Sub

Private Sub Wsinviofile_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "non e' possibile effettuare la connessione al server....."
End Sub

Private Sub Wsinviofile_Connect()
 avviso.Show
 avviso.Labelmessaggio.Caption = " la connessione e' stabilita "
End Sub


Private Sub Wsinviofile_DataArrival(ByVal bytesTotal As Long)
Dim strData As String ' received datas'
Dim strBuffer As String 'Buffer'

Wsinviofile.GetData strData 'get data
'if there are more than 2048 left
If FileLen(DlgSend.FileName) - LngCursor < 2048 Then
    'Buffer size adjustment
    strBuffer = Space(FileLen(DlgSend.FileName) - LngCursor)
'if there are more than 2048 bytes left
ElseIf FileLen(DlgSend.FileName) - LngCursor > 2048 Then
    'buffer = 2048
    strBuffer = Space(2048)
End If
'if pointer value = source file size
If FileLen(DlgSend.FileName) = LngCursor Then
    'we have finished the transfert, we close the opened file, and we tell the server the job is done
    LblStatut.Caption = "Statut: " & DlgSend.FileName & " successfully uploaded"
    Wsinviofile.SendData "E"
    Close #1
    Exit Sub
End If
'the server ask for the transfert beginning
If Left(strData, 1) = "S" Then
    LblStatut.Caption = "Statut: Uploading " & DlgSend.FileName
    'we open the file in binary mode
    Open DlgSend.FileName For Binary As #1
        Get #1, , strBuffer
    Wsinviofile.SendData strBuffer ' we send it'
    LngCursor = Len(strBuffer) ' at the beginning of a transfert it's 2048'
'The server ask for another file chunk
ElseIf Left(strData, 1) = "N" Then
        Get #1, LngCursor + 1, strBuffer
    Wsinviofile.SendData strBuffer 'we send'
    LngCursor = LngCursor + Len(strBuffer) 'we update the pointer'
End If
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
 Cancel = 1
 Cmddisconnect_Click
End Sub
