VERSION 5.00
Begin VB.Form listaamici 
   Caption         =   "lista amici"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdload 
      Caption         =   "carica lista amici"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Cmdsalva 
      Caption         =   "salva lista amici"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Cmdesci 
      Caption         =   "esci"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Ccancella 
      Caption         =   "cancella lista amici"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Cmdrimuovi 
      Caption         =   "rimuovi"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Txtrecordrimuovi 
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Txtrecord 
      BackColor       =   &H80000009&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ListBox Listusers 
      Height          =   2985
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Cmdamici 
      Caption         =   "amici >>>>"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox Listfriends 
      Height          =   2985
      Left            =   4800
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Labelcancella 
      BackStyle       =   0  'Transparent
      Caption         =   "cancella tutta la lista amici"
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "seleziona un contatto dalla lista amici per rimuoverlo"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Labelamici 
      BackStyle       =   0  'Transparent
      Caption         =   "lista amici"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Labelutenti 
      BackStyle       =   0  'Transparent
      Caption         =   "lista utenti"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "listaamici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdload_Click()
Dim MyString As String
    On Error Resume Next

    Open App.Path & "\log\liste\amici\savelist.dat" For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
       If Len(MyString$) = 0 Or MyString$ = " " Then
       
       GoTo 10
       End If
       
        Listfriends.AddItem MyString$
        Listfriends.Refresh
10
    Wend
20
    Close #1
End Sub

Private Sub Cmdsalva_Click()
Dim savelist As Long
    On Error Resume Next
    Open App.Path & "\log\liste\amici\savelist.dat" For Output As #1
    For savelist& = 0 To Listfriends.ListCount - 1
        Print #1, Listfriends.List(savelist&)
    Next savelist&
    Close #1
End Sub

Private Sub form_load()
Dim X 'dim x'

For X = 0 To chat.Listusers.ListCount - 1

listaamici.Listusers.AddItem chat.Listusers.List(X) 'trasferiamo i dati da una lista ad unaltra'

Next X

Cmdload_Click
End Sub
Private Sub Ccancella_Click()
Listfriends.Clear
End Sub

Private Sub Cmdamici_Click()
Listfriends.AddItem Txtrecord.Text
End Sub

Private Sub Cmdesci_Click()
listaamici.Visible = False
End Sub

Private Sub Cmdrimuovi_Click()
Dim i As Integer 'dichiariamo la variabile'
For i = (Listfriends.ListCount - 1) To 0 Step -1

      If InStr(Listfriends.List(i), Txtrecordrimuovi.Text) <> 0 Then 'mettiamo in un txt apparte l'utente selezionato'
      Listfriends.RemoveItem i 'rimuoviamo l'utente'
      End If
  Next
End Sub

Private Sub Listfriends_Click()
Txtrecordrimuovi.Text = Listfriends.Text
End Sub

Private Sub Listusers_Click()
Txtrecord.Text = Listusers.Text
End Sub
 
 Private Sub form_unload(cancel As Integer)
 cancel = 1
 Cmdsalva_Click
 Cmdesci_Click
 End Sub
