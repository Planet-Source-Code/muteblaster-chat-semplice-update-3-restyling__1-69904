VERSION 5.00
Begin VB.Form avviso_connessione 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000013&
         Caption         =   "non mostrare in futuro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "connessione con il server della chat avvenuta con successo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "e' in corso la connessione con il server della chat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AUTENTICAZIONE AVVENUTA CON SUCCESSO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "e' stata stabilita la connessione con il  server del messenger"
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
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
   End
End
Attribute VB_Name = "avviso_connessione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub formload()
 If opzioni.Check1 = 1 Then
    avviso_connessione.Visible = True
 End If
End Sub

Private Sub Check1_Click()
 If Check1 = 1 Then
    opzioni.Check1 = 1
 End If
End Sub
