VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form database_immagini_personali 
   Caption         =   "archivio immagini personali"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4185
   LinkTopic       =   "Form2"
   ScaleHeight     =   2820
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DataField       =   "Pic"
      DataSource      =   "Data1"
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Immagine"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programmi\Microsoft Visual Basic\ImgData\dbimg.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prova"
      Top             =   2400
      Width           =   2652
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apri"
      Default         =   -1  'True
      Height          =   516
      Left            =   2868
      TabIndex        =   2
      Top             =   0
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Esci"
      Height          =   492
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1212
   End
   Begin VB.CommandButton cmdrimuovi 
      Caption         =   "rimuovi"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   2880
      Top             =   2160
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
End
Attribute VB_Name = "database_immagini_personali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdrimuovi_Click()
 Data1.Recordset.Delete
 Data1.Recordset.MoveNext
End Sub

Private Sub Command1_Click()
CMDialog1.Filename = ""
CMDialog1.Action = 1
If CMDialog1.Filename <> "" Then
  Data1.Recordset.AddNew 'crea nuovo spazio
  Picture1.Picture = LoadPicture(CMDialog1.Filename) 'carica l'immagine
  Data1.Recordset.Update 'salva
  Data1.Recordset.MoveLast 'sposta sull'immagine appena immessa
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 Data1.DatabaseName = App.Path & "\database immagini" & "\personali" & "\dbimg.mdb"
End Sub

