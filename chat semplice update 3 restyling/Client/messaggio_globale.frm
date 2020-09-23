VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form messaggio_globale 
   Caption         =   "global message"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtmessage 
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"messaggio_globale.frx":0000
   End
End
Attribute VB_Name = "messaggio_globale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
