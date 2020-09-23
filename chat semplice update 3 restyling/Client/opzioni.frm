VERSION 5.00
Begin VB.Form opzioni 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "opzioni"
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Height          =   2055
      Left            =   5760
      TabIndex        =   15
      Top             =   2880
      Width           =   3015
      Begin client.CandyButton Cmdbiglietto 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "modifica biglietto da visita"
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
   End
   Begin VB.Timer Timer_salvataggio 
      Interval        =   60000
      Left            =   5160
      Top             =   4920
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   2055
      Left            =   0
      TabIndex        =   11
      Top             =   2880
      Width           =   5655
      Begin VB.CheckBox Check5 
         BackColor       =   &H80000013&
         Caption         =   "blocca popup"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H80000013&
         Caption         =   "all'avvio proteggi con password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000013&
         Caption         =   "disattiva biglietto da visita"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000013&
         Caption         =   "mostra introduzione"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000013&
         Caption         =   "avviso connessione"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   2400
         X2              =   2400
         Y1              =   120
         Y2              =   2040
      End
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H80000013&
      Caption         =   "ok"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   2775
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Cmdriabilita_extrachat 
         BackColor       =   &H80000013&
         Caption         =   "riabilita extrachat"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdrispostefatte 
         BackColor       =   &H80000013&
         Caption         =   "risposte fatte"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcriptafile 
         BackColor       =   &H80000013&
         Caption         =   "Cripta file"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Cmddecriptafile 
         BackColor       =   &H80000013&
         Caption         =   "decripta file"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton CmdGNU 
         BackColor       =   &H80000013&
         Caption         =   "licenza"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.Frame Framelog 
      BackColor       =   &H80000013&
      Caption         =   "opzioni log chat"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Cmdstampa 
         BackColor       =   &H80000013&
         Caption         =   "stampa"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton mdsalva 
         BackColor       =   &H80000013&
         Caption         =   "salva"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcancella 
         BackColor       =   &H80000013&
         Caption         =   "cancella"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "opzioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Sub Check4_Click()
' son costretto a mettere la condizione che il login non sia visibile'
' per colpa di uno bug che all'avvio il form psw_sicurezza si caricava comunque'
' cosi' lo ho arginato, infatti se il form login e' visibile non parte il form'
' psw_sicurezza'.....
'-------------
' 10 dicembre 2007 scoperto il bug
' quando si caricano le impostazioni icheck e' come se venissero ripremuti, quindi se quando premi un check
' fai visualizzare un form , al caricamento delle impostazioni e' come se venisse ripremuto....ecco perche'
' partiva il form.....

If opzioni.Visible = True Then
 If Check4 = 1 Then
  psw_sicurezza.Show
 End If
End If
End Sub

Private Sub Check5_Click()
If Check5 = 1 Then
   popup.Visible = False
End If
End Sub

Private Sub Cmdbiglietto_Click()
 biglietto_da_visita.Show
End Sub

Private Sub Form_Load()
 Check1.Value = RegLoad(Check1)
 Check2.Value = RegLoad(Check2)
 Check4.Value = RegLoad(Check4)
 Check5.Value = RegLoad(Check5)
End Sub


Private Sub SaveControlValues()
 Call RegSave(Check1, Check1.Value)
 Call RegSave(Check2, Check2.Value)
 Call RegSave(Check4, Check4.Value)
 Call RegSave(Check5, Check5.Value)
End Sub

Private Sub Cmdcancella_Click()
chat.txtChat.Text = ""
End Sub

Private Sub Cmdcriptafile_Click()
criptafile.Show
End Sub

Private Sub Cmddecriptafile_Click()
decriptafile.Show
End Sub

Private Sub CmdGNU_Click()
 licenza.Show
End Sub

Private Sub Cmdok_Click()
 Call SaveControlValues
 chat.Picture6.Top = 10800
End Sub

Private Sub Cmdriabilita_extrachat_Click()
riabilita_privat.Show
End Sub

Private Sub Cmdrispostefatte_Click()
 risposte.Show
End Sub

Private Sub Cmdstampa_Click()
 On Error GoTo PrintErr
    chat.Cmdl.Flags = cdlPDHidePrintToFile + cdlPDNoPageNums
   chat.Cmdl.ShowPrinter
    Printer.ScaleLeft = -((Printer.Width - chat.txtChat.Width) / 2)
    Printer.Print chat.txtChat.Text
    Printer.EndDoc
PrintErr:
End Sub

Private Sub mdsalva_Click()
On Error GoTo SaveErr
   chat.Cmdl.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
   chat.Cmdl.Filter = "Text Files (*.txt)|*.txt"
    chat.Cmdl.ShowSave
    Open chat.Cmdl.Filename For Output As #1
    Print #1, chat.txtChat.Text 'saves file
    Close #1
SaveErr:
End Sub

 Private Sub Form_Unload(Cancel As Integer)
 Cancel = 1
 Cmdok_Click
 End Sub
 

Private Sub Timer_salvataggio_Timer()
 Cmdok_Click
End Sub
