VERSION 5.00
Begin VB.Form frmDeleteUsr 
   Caption         =   "Delete User"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmDeleteUsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = '" & txtUser.Text & "'")
    
    If rs.RecordCount > 0 Then
        rs.Delete
        
        If Dir("buddylists\" & txtUser & ".txt") <> "" Then
            Open "buddylists\" & txtUser & ".txt" For Input As #1
            
            Line Input #1, numusers
            
            For i = 1 To numusers
                Line Input #1, auser
                Open "buddyref\" & auser & ".txt" For Input As #2
                
                Line Input #2, numusers2
                
                For j = 1 To numusers2
                    Line Input #2, buser
                    If buser <> txtUser Then
                        buddyref = buddyref & buser & vbCrLf
                    End If
                Next
                Close #2
                Open "buddyref\" & auser & ".txt" For Output As #2
                Write #2, numusers2 - 1
                Print #2, Left(buddyref, Len(buddyref) - 1)
                Close #2
            Next
            Close #1
        End If
        
        Kill ("buddylists\" & txtUser & ".txt")
        Kill ("buddyref\" & txtUser & ".txt")
        Kill ("buddyinfo\" & txtUser & ".txt")
    Else
        MsgBox "User doesn't exist!", , "Error!"
    End If
    txtUser.Text = ""
    MsgBox "Deleted.", , "Success!"
End Sub
