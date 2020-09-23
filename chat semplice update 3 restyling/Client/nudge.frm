VERSION 5.00
Begin VB.Form nudge 
   Caption         =   "nudge"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   1980
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timernudge_chat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   720
   End
   Begin VB.Timer TimernudgeMSricevi 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   720
   End
   Begin VB.CommandButton Cmdnudge 
      Caption         =   "nudge"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Timer Timernudge1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   720
   End
End
Attribute VB_Name = "nudge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------------------------------------------'
'            inizio esperimento del codice per il nudge                                 '
' ma non sono mica tanto convinto che sia una buona idea                                '
'---------------------------------------------------------------------------------------'
Public Flg1 As Integer
Public FTOP As Integer
Public FLEFT As Integer

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H1



Private Sub form_load()
Flg1 = 0
End Sub

 ' ---------------INIZIO ESPERIMENTO NUDGE ------------------------ '
 
Private Sub Cmdnudge_Click()
Timernudge1.Enabled = True
soundfile$ = App.Path & "\msinvio.wav" ' Plays sound '
wflags% = SND_ASYNC Or SND_NODEFAULT
HaHa = sndPlaySound(soundfile$, wflags%)
End Sub

Private Sub Timernudge_chat_Timer()
Select Case Flg1 ' muoviamo il form grazie al timer ....ci vuole una vita '
Case 0           ' a fare sto' codice ma sicuramente questa non e' la soluzione piu' semplice :)'
FTOP = chat.Top
FLEFT = chat.Left
chat.Left = chat.Left + 30
chat.Top = chat.Top + 30
Flg1 = Flg1 + 1
Case 1
chat.Left = chat.Left - 45
chat.Top = chat.Top - 45
Flg1 = Flg1 + 1
Case 2
chat.Left = chat.Left + 60
chat.Top = chat.Top + 60
Flg1 = Flg1 + 1
Case 3
chat.Left = chat.Left - 75
chat.Top = chat.Top - 75
Flg1 = Flg1 + 1
Case 4
chat.Left = chat.Left + 90
chat.Top = chat.Top + 90
Flg1 = Flg1 + 1
Case 5
chat.Left = chat.Left - 105
chat.Top = chat.Top - 105
Flg1 = Flg1 + 1
Case 6
chat.Left = chat.Left + 105
chat.Top = chat.Top + 105
Flg1 = Flg1 + 1
Case 7
chat.Left = chat.Left - 75
chat.Top = chat.Top - 75
Flg1 = Flg1 + 1
Case 8
chat.Left = chat.Left + 90
chat.Top = chat.Top + 90
Flg1 = Flg1 + 1
Case 9
chat.Left = chat.Left - 135
chat.Top = chat.Top - 135
Flg1 = Flg1 + 1
Case 10
chat.Left = chat.Left + 90
chat.Top = chat.Top + 90
Flg1 = Flg1 + 1
Case 11
chat.Left = chat.Left - 105
chat.Top = chat.Top - 105
Flg1 = Flg1 + 1
Case 12
chat.Left = chat.Left + 135
chat.Top = chat.Top + 135
Flg1 = Flg1 + 1
Case 13
chat.Left = chat.Left - 90
chat.Top = chat.Top - 90
Flg1 = Flg1 + 1
Case 14
chat.Left = chat.Left + 75
chat.Top = chat.Top + 75
Flg1 = Flg1 + 1
Case 15
chat.Left = chat.Left - 150
chat.Top = chat.Top - 150
Flg1 = Flg1 + 1
Case 16
chat.Left = chat.Left + 105
chat.Top = chat.Top + 105
Flg1 = Flg1 + 1
Case 17
chat.Left = chat.Left - 75
chat.Top = chat.Top - 75
Flg1 = Flg1 + 1
Case 18
chat.Left = chat.Left + 90
chat.Top = chat.Top + 90
Flg1 = Flg1 + 1
Case 19
chat.Left = chat.Left - 105
chat.Top = chat.Top - 105
Flg1 = Flg1 + 1
Case 20
chat.Left = chat.Left + 135
chat.Top = chat.Top + 135
Flg1 = Flg1 + 1
Case 21
chat.Left = chat.Left - 150
chat.Top = chat.Top - 150
Flg1 = Flg1 + 1
Case 22
chat.Left = chat.Left + 180
chat.Top = chat.Top + 180
Flg1 = Flg1 + 1
Case 23
chat.Left = chat.Left - 150
chat.Top = chat.Top - 150
Flg1 = Flg1 + 1
Case 24
chat.Left = chat.Left + 195
chat.Top = chat.Top + 195
Flg1 = Flg1 + 1
Case 25
chat.Left = FLEFT
chat.Top = FTOP
Flg1 = 0
Timernudge_chat.Enabled = False
End Select
End Sub

Private Sub Timernudge1_Timer()
Select Case Flg1 ' muoviamo il form grazie al timer ....ci vuole una vita '
Case 0           ' a fare sto' codice ma sicuramente questa non e' la soluzione piu' semplice :)'
FTOP = MSinvio.Top
FLEFT = MSinvio.Left
MSinvio.Left = MSinvio.Left + 30
MSinvio.Top = MSinvio.Top + 30
Flg1 = Flg1 + 1
Case 1
MSinvio.Left = MSinvio.Left - 45
MSinvio.Top = MSinvio.Top - 45
Flg1 = Flg1 + 1
Case 2
MSinvio.Left = MSinvio.Left + 60
MSinvio.Top = MSinvio.Top + 60
Flg1 = Flg1 + 1
Case 3
MSinvio.Left = MSinvio.Left - 75
MSinvio.Top = MSinvio.Top - 75
Flg1 = Flg1 + 1
Case 4
MSinvio.Left = MSinvio.Left + 90
MSinvio.Top = MSinvio.Top + 90
Flg1 = Flg1 + 1
Case 5
MSinvio.Left = MSinvio.Left - 105
MSinvio.Top = MSinvio.Top - 105
Flg1 = Flg1 + 1
Case 6
MSinvio.Left = MSinvio.Left + 105
MSinvio.Top = MSinvio.Top + 105
Flg1 = Flg1 + 1
Case 7
MSinvio.Left = MSinvio.Left - 75
MSinvio.Top = MSinvio.Top - 75
Flg1 = Flg1 + 1
Case 8
MSinvio.Left = MSinvio.Left + 90
MSinvio.Top = MSinvio.Top + 90
Flg1 = Flg1 + 1
Case 9
MSinvio.Left = MSinvio.Left - 135
MSinvio.Top = MSinvio.Top - 135
Flg1 = Flg1 + 1
Case 10
MSinvio.Left = MSinvio.Left + 90
MSinvio.Top = MSinvio.Top + 90
Flg1 = Flg1 + 1
Case 11
MSinvio.Left = MSinvio.Left - 105
MSinvio.Top = MSinvio.Top - 105
Flg1 = Flg1 + 1
Case 12
MSinvio.Left = MSinvio.Left + 135
MSinvio.Top = MSinvio.Top + 135
Flg1 = Flg1 + 1
Case 13
MSinvio.Left = MSinvio.Left - 90
MSinvio.Top = MSinvio.Top - 90
Flg1 = Flg1 + 1
Case 14
MSinvio.Left = MSinvio.Left + 75
MSinvio.Top = MSinvio.Top + 75
Flg1 = Flg1 + 1
Case 15
MSinvio.Left = MSinvio.Left - 150
MSinvio.Top = MSinvio.Top - 150
Flg1 = Flg1 + 1
Case 16
MSinvio.Left = MSinvio.Left + 105
MSinvio.Top = MSinvio.Top + 105
Flg1 = Flg1 + 1
Case 17
MSinvio.Left = MSinvio.Left - 75
MSinvio.Top = MSinvio.Top - 75
Flg1 = Flg1 + 1
Case 18
MSinvio.Left = MSinvio.Left + 90
MSinvio.Top = MSinvio.Top + 90
Flg1 = Flg1 + 1
Case 19
MSinvio.Left = MSinvio.Left - 105
MSinvio.Top = MSinvio.Top - 105
Flg1 = Flg1 + 1
Case 20
MSinvio.Left = MSinvio.Left + 135
MSinvio.Top = MSinvio.Top + 135
Flg1 = Flg1 + 1
Case 21
MSinvio.Left = MSinvio.Left - 150
MSinvio.Top = MSinvio.Top - 150
Flg1 = Flg1 + 1
Case 22
MSinvio.Left = MSinvio.Left + 180
MSinvio.Top = MSinvio.Top + 180
Flg1 = Flg1 + 1
Case 23
MSinvio.Left = MSinvio.Left - 150
MSinvio.Top = MSinvio.Top - 150
Flg1 = Flg1 + 1
Case 24
MSinvio.Left = MSinvio.Left + 195
MSinvio.Top = MSinvio.Top + 195
Flg1 = Flg1 + 1
Case 25
MSinvio.Left = FLEFT
MSinvio.Top = FTOP
Flg1 = 0
Timernudge1.Enabled = False
End Select
End Sub


Private Sub TimernudgeMSricevi_Timer()
Select Case Flg1 ' muoviamo il form grazie al timer ....ci vuole una vita '
Case 0           ' a fare sto' codice ma sicuramente questa non e' la soluzione piu' semplice :)'
FTOP = MSricevi.Top
FLEFT = MSricevi.Left
MSricevi.Left = MSricevi.Left + 30
MSricevi.Top = MSricevi.Top + 30
Flg1 = Flg1 + 1
Case 1
MSricevi.Left = MSricevi.Left - 45
MSricevi.Top = MSricevi.Top - 45
Flg1 = Flg1 + 1
Case 2
MSricevi.Left = MSricevi.Left + 60
MSricevi.Top = MSricevi.Top + 60
Flg1 = Flg1 + 1
Case 3
MSricevi.Left = MSricevi.Left - 75
MSricevi.Top = MSricevi.Top - 75
Flg1 = Flg1 + 1
Case 4
MSricevi.Left = MSricevi.Left + 90
MSricevi.Top = MSricevi.Top + 90
Flg1 = Flg1 + 1
Case 5
MSricevi.Left = MSricevi.Left - 105
MSricevi.Top = MSricevi.Top - 105
Flg1 = Flg1 + 1
Case 6
MSricevi.Left = MSricevi.Left + 105
MSricevi.Top = MSricevi.Top + 105
Flg1 = Flg1 + 1
Case 7
MSricevi.Left = MSricevi.Left - 75
MSricevi.Top = MSricevi.Top - 75
Flg1 = Flg1 + 1
Case 8
MSricevi.Left = MSricevi.Left + 90
MSricevi.Top = MSricevi.Top + 90
Flg1 = Flg1 + 1
Case 9
MSricevi.Left = MSricevi.Left - 135
MSricevi.Top = MSricevi.Top - 135
Flg1 = Flg1 + 1
Case 10
MSricevi.Left = MSricevi.Left + 90
MSricevi.Top = MSricevi.Top + 90
Flg1 = Flg1 + 1
Case 11
MSricevi.Left = MSricevi.Left - 105
MSricevi.Top = MSricevi.Top - 105
Flg1 = Flg1 + 1
Case 12
MSricevi.Left = MSricevi.Left + 135
MSricevi.Top = MSricevi.Top + 135
Flg1 = Flg1 + 1
Case 13
MSricevi.Left = MSricevi.Left - 90
MSricevi.Top = MSricevi.Top - 90
Flg1 = Flg1 + 1
Case 14
MSricevi.Left = MSricevi.Left + 75
MSricevi.Top = MSricevi.Top + 75
Flg1 = Flg1 + 1
Case 15
MSricevi.Left = MSricevi.Left - 150
MSricevi.Top = MSricevi.Top - 150
Flg1 = Flg1 + 1
Case 16
MSricevi.Left = MSricevi.Left + 105
MSricevi.Top = MSricevi.Top + 105
Flg1 = Flg1 + 1
Case 17
MSricevi.Left = MSricevi.Left - 75
MSricevi.Top = MSricevi.Top - 75
Flg1 = Flg1 + 1
Case 18
MSricevi.Left = MSricevi.Left + 90
MSricevi.Top = MSricevi.Top + 90
Flg1 = Flg1 + 1
Case 19
MSricevi.Left = MSricevi.Left - 105
MSricevi.Top = MSricevi.Top - 105
Flg1 = Flg1 + 1
Case 20
MSricevi.Left = MSricevi.Left + 135
MSricevi.Top = MSricevi.Top + 135
Flg1 = Flg1 + 1
Case 21
MSricevi.Left = MSricevi.Left - 150
MSricevi.Top = MSricevi.Top - 150
Flg1 = Flg1 + 1
Case 22
MSricevi.Left = MSricevi.Left + 180
MSricevi.Top = MSricevi.Top + 180
Flg1 = Flg1 + 1
Case 23
MSricevi.Left = MSricevi.Left - 150
MSricevi.Top = MSricevi.Top - 150
Flg1 = Flg1 + 1
Case 24
MSricevi.Left = MSricevi.Left + 195
MSricevi.Top = MSricevi.Top + 195
Flg1 = Flg1 + 1
Case 25
MSricevi.Left = FLEFT
MSricevi.Top = FTOP
Flg1 = 0
TimernudgeMSricevi.Enabled = False
End Select
End Sub

' -------------------FINE ESPERIMENTO NUDGE --------- '
