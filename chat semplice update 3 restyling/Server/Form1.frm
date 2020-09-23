VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmServer 
   AutoRedraw      =   -1  'True
   Caption         =   "Server"
   ClientHeight    =   8595
   ClientLeft      =   255
   ClientTop       =   645
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12555
   Begin VB.Timer Timer_setparent 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7800
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "server messenger"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdDeleteUsr"
      Tab(0).Control(1)=   "txtLog"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "server chat"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "server multichat rooms"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "server chat moderatori"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture3"
      Tab(3).ControlCount=   1
      Begin VB.PictureBox Picture3 
         Height          =   7455
         Left            =   -74880
         ScaleHeight     =   7395
         ScaleWidth      =   11955
         TabIndex        =   5
         Top             =   840
         Width           =   12015
      End
      Begin VB.PictureBox Picture2 
         Height          =   7335
         Left            =   -74760
         ScaleHeight     =   7275
         ScaleWidth      =   11835
         TabIndex        =   4
         Top             =   720
         Width           =   11895
      End
      Begin VB.PictureBox Picture1 
         Height          =   7455
         Left            =   120
         ScaleHeight     =   7395
         ScaleWidth      =   11955
         TabIndex        =   3
         Top             =   720
         Width           =   12015
      End
      Begin VB.CommandButton cmdDeleteUsr 
         Caption         =   "Delete User"
         Height          =   495
         Left            =   -68280
         TabIndex        =   2
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtLog 
         Height          =   6375
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1020
         Width           =   6255
      End
   End
   Begin MSWinsockLib.Winsock win 
      Index           =   0
      Left            =   3840
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Dim Username(1 To 49) As String
Dim Password(1 To 49) As String

Private Sub cmdDeleteUsr_Click()
    frmDeleteUsr.Show
End Sub

Private Sub Form_Load()
   'all'avvio il timer set parent distribuita' tutti form nei tab'
   Timer_setparent.Enabled = True
    txtLog.Text = "Loading server..." & vbCrLf
    'This code loads all the winsock controls automatically into the 'win' array
    For I = 1 To 49
        Load win(I)
    Next
    'Listens on port 9012584 for incoming connections
    win(0).LocalPort = 12584
    win(0).Listen
    txtLog.Text = txtLog.Text & "Server running..." & vbCrLf & vbCrLf
    
    'This connects to the database of users
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase("Data.mdb")
    
    'Set's all user status to offline. If you shut down the server and people
    'Are online, they won't be signed off in the database so this code does that.
    Set rs = db.OpenRecordset("SELECT * FROM Users")
    Do While Not rs.EOF
        rs.Edit
        rs!Online = 0
        rs.Update
        rs.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer_setparent_Timer()
    SetParent server.hWnd, frmServer.Picture1.hWnd
    server.Show
    server.Move 0, 0
    SetParent server_multicanale.hWnd, frmServer.Picture2.hWnd
    server_multicanale.Show
    server_multicanale.Move 0, 0
    SetParent server_MOD_chat.hWnd, frmServer.Picture3.hWnd
    server_MOD_chat.Show
    server_MOD_chat.Move 0, 0
    Timer_setparent.Enabled = False
End Sub

Private Sub txtLog_Change()
    txtLog.SelStart = Len(txtLog)
End Sub

Private Sub win_Close(Index As Integer)
    'Set's user offline in database
    Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = '" & Username(Index) & "'")
    If rs.RecordCount > 0 Then
        rs.Edit
        rs!Online = 0
        rs.Update
    End If
    txtLog = txtLog & Username(Index) & " logged out." & vbCrLf
    txtLog = txtLog & "Lost connection to " & win(Index).RemoteHostIP & vbCrLf
    'My buddy ref system.
    'Every time you add someone to your buddy list, your username
    'get's added to the added user's Reference list...The reference list just tells the
    'server who to inform when someone signs on or off.
    If Dir("buddyref\" & Username(Index) & ".txt") <> "" Then
        Close (Index)
        Open "buddyref\" & Username(Index) & ".txt" For Input As #Index
        Line Input #Index, numusers
        
        'This goes through all the connections to see if the user that needs to be
        'informed that his buddy signed on or off, is online
        For I = 1 To numusers
            Line Input #Index, suser
            For j = 1 To 49
                If UCase(Username(j)) = UCase(suser) Then
                    If win(j).State = 7 Then
                        win(j).SendData "ss-" & Username(Index)
                    End If
                    
                    j = 50
                End If
                DoEvents
            Next
            DoEvents
        Next
    End If
    'Set's the username associated with this connection to nothing and closes
    'the connection
    Username(Index) = ""
    Close (Index)
    win(Index).Close
End Sub

Private Sub win_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'Goes through all winsock controls to see if any are free and if one is free,
    'the program uses it to establish a connection with the user
    For I = 1 To 49
        If win(I).State = 0 Or win(I).State <> 7 Then
            win(I).Close
            win(I).Accept requestID
            txtLog.Text = txtLog.Text & "Connected to " & win(I).RemoteHostIP & vbCrLf
            I = 50
        End If
        DoEvents
    Next
End Sub

Private Sub win_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Buffer As String
    Dim msg As String
    win(Index).GetData msg, vbString
    Buffer = msg

    'This shit checks for two messages coming in at the same time
    nummessages = 0
    If InStr(1, msg, "\-") > 0 Then
    If Len(msg) > InStr(1, msg, "\-") + 1 Then
        For I = 1 To Len(msg) - 1
            If Mid(msg, I, 2) = "\-" Then
                nummessages = nummessages + 1
            End If
        Next
    Else
        'msg = Replace(msg, "\-", "")
        nummessages = 1
    End If
    Else
    nummessages = 1
    End If
    For q = 1 To nummessages
        If InStr(1, Buffer, "\-") > 0 Then
            msg = Mid(Buffer, 1, InStr(1, Buffer, "\-") - 1)
            Buffer = Mid(Buffer, InStr(1, Buffer, "\-") + 2)
        End If
        
    'The following code checks for the Username and the Password
        If Left(msg, 5) = "User-" Then
            'This extracts the Username from the message
            Username(Index) = Mid(msg, 6)
        ElseIf Left(msg, 5) = "Pass-" Then
            'This extracts the Password from the message
            Password(Index) = Mid(msg, 6)
            'This code actually compares the password gotten from the user to the password
            'in the database
            Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = '" & Username(Index) & "'")
            If rs.RecordCount > 0 Then
                If rs.Fields(2) = 0 Then
                    If Password(Index) = rs.Fields(1) Then
                        'This code is executed after the Username and Password is verified
                        'This person has the right info
                        win(Index).SendData "Login Success\-"
                        win(Index).SendData "forsn-" & rs.Fields(0) & "\-"
                        Username(Index) = rs.Fields(0)
                        txtLog.Text = txtLog.Text & Username(Index) & " logged in." & vbCrLf
                        'This updates the users Online record in the database
                        rs.Edit
                        rs!Online = 1
                        rs.Update
                        'This code is cool. Checks the userse reference file for a list of
                        'all the people that have this user in their buddy list, and tells them
                        'he's signing on.
                        If Dir("buddyref/" & Username(Index) & ".txt") <> "" Then
                            filenum = FreeFile
                            Open "buddyref/" & Username(Index) & ".txt" For Input As #filenum
                            Line Input #filenum, numusers
                            For I = 1 To numusers
                                Line Input #filenum, suser
                                If suser <> Username(Index) Then
                                    For j = 1 To 49
                                        If UCase(Username(j)) = UCase(suser) Then
                                            'If win(j).State = 7 Then
                                                'This sends the message to the person who has the
                                                'logging-on user in their buddy list and tells their
                                                'buddy list to show the person as "online"
                                                win(j).SendData "ss-" & Username(Index)
                                                
                                            End If
                                       ' End If
                                        DoEvents
                                    Next
                                End If
                                DoEvents
                            Next
                            Close (filenum)
                        End If
                        filenum = FreeFile
                        'This sends the user his or her own profile so he/she can edit it.
                        If Dir("buddyinfo\" & Username(Index) & ".txt") <> "" Then
                            filenum = FreeFile
                            Open "buddyinfo\" & Username(Index) & ".txt" For Input As #filenum
                            
                            Do While Not EOF(filenum)
                                Line Input #filenum, temp
                                info = info & temp & vbCrLf
                            Loop
                            Close (filenum)
                            win(Index).SendData "ginf-" & Username(Index) & "-" & info & "\-"
                        Else
                            win(Index).SendData "ginf-" & Username(Index) & "-\-"
                        End If
                    Else
                        'Wrong password or username given'
                        win(Index).SendData "msg-Error! Wrong password!"
                    End If
                Else
                'User logged on already!
                win(Index).SendData "msg-Error! User already logged in!"
            End If
        Else
            'Wrong password
            win(Index).SendData "msg-Error! User doesn't exist!"
        End If
                'Someones sending an IM.
        ElseIf Left(msg, 3) = "im-" Then
            'This code extracts the recievers username and the actual message being sent.
            pos = InStr(4, msg, "-")
            suser = Mid(msg, 4, (pos) - 4)
            im = Mid(msg, pos + 1)
    
            found = False
            'This checks to see if the recieving person is online. If we find a socket
            'that is connected to the recieving user, we send the message...if not, we
            'tell the sender the person is offline or doesnt exist.
            For I = 1 To 49
                If UCase(Username(I)) = UCase(suser) Then
                    win(I).SendData "im-" & Username(Index) & "-" & im
                    found = True
                    I = 50
                End If
                DoEvents
            Next
            If found = False Then
                win(Index).SendData "msg-User offline or doesn't exist!"
               
            End If
            
        'Once the password was verified, starts sending the user's buddylist
        ElseIf Left(msg, 8) = "Send b/l" Then
            If Dir("buddylists/" & Username(Index) & ".txt") <> "" Then
                filenum2 = FreeFile
                Open "buddylists/" & Username(Index) & ".txt" For Input As #filenum2
                
                Line Input #filenum2, numusers
                For I = 1 To numusers
                    Line Input #filenum2, nuser
                    Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = """ & nuser & """")
                    If rs.RecordCount > 0 Then
                        Online = rs.Fields(2)
                    Else
                        Online = False
                    End If
                    'If the user on the buddy list is online, it tells the client
                    'to add the user to the online portion of the buddy list or
                    'vice versa
                    If (Online = 1) Then
                        win(Index).SendData "cusr-" & nuser & "\-"
                       
                    Else
                        win(Index).SendData "dusr-" & nuser & "\-"
                        
                    End If
                    DoEvents
                Next
                Close (filenum2)
            End If
            'Tells the client the server is done sending the buddy list.
            win(Index).SendData "End b/l\-"
            'This adds a buddy to the users buddy list
         ElseIf Left(msg, 4) = "add-" Then
            'This is used to get the correct formatting of the screenname as it is in the DB
            Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = """ & Mid(msg, 5) & """")
            If rs.RecordCount > 0 Then
                'File manipulation, it just adds the user to the end of the buddy list
                If Dir("buddylists/" & Username(Index) & ".txt") <> "" Then
                    filenum3 = FreeFile
                    Open "buddylists/" & Username(Index) & ".txt" For Input As #filenum3
                    Line Input #filenum3, numusers
                    For I = 1 To numusers
                        Line Input #filenum3, temp
                        buddylist = buddylist & temp & vbCrLf
                        DoEvents
                    Next
                    Close (filenum3)
                End If
                filenum3 = FreeFile
                Open "buddylists/" & Username(Index) & ".txt" For Output As #filenum3
                buddylist = buddylist & rs.Fields(0)
                Write #filenum3, (numusers + 1)
                Print #filenum3, buddylist
                Close (filenum3)
                
                'Adds user to clients buddy list as either on or offline
                If rs.Fields(2) = 1 Then
                    win(Index).SendData "cusr-" & rs.Fields(0)
                Else
                    win(Index).SendData "dusr-" & rs.Fields(0)
                End If
                
                
                numusers = 0
                buddylist = ""
                'Adds the user who's adding the buddy to the buddy's reference buddy list
                If Dir("buddyref/" & Mid(msg, 5) & ".txt") <> "" Then
                    filenum3 = FreeFile
                    Open "buddyref/" & Mid(msg, 5) & ".txt" For Input As #filenum3
                    Line Input #filenum3, numusers
                    For I = 1 To numusers
                        Line Input #filenum3, temp
                        buddylist = buddylist & temp & vbCrLf
                        DoEvents
                    Next
                    Close (filenum3)
                End If
                Open "buddyref/" & Mid(msg, 5) & ".txt" For Output As #Index
                buddylist = buddylist & Username(Index)
                Write #Index, numusers + 1
                Print #Index, buddylist
                Close (Index)
            Else
                win(Index).SendData "msg-User doesn't exist!"
            End If
        'Deletes someone from user's buddylist
        ElseIf Left(msg, 4) = "del-" Then
            filenum4 = FreeFile
            Open "buddylists/" & Username(Index) & ".txt" For Input As #filenum4
            Line Input #filenum4, numusers
            For I = 1 To numusers
                todel = Mid(msg, 5)
                Line Input #filenum4, temp
                If (temp <> todel) Then
                    buddylist = buddylist & temp & vbCrLf
                End If
                DoEvents
            Next
            Close (filenum4)
            filenum4 = FreeFile
            Open "buddylists/" & Username(Index) & ".txt" For Output As #filenum4
            Write #filenum4, numusers - 1
            If numusers - 1 > 0 Then
                Print #filenum4, Left(buddylist, Len(buddylist) - 1)
            End If
            Close (filenum4)
            
            buddylist = ""
            If Dir("buddyref/" & Mid(msg, 5) & ".txt") <> "" Then
                    filenum4 = FreeFile
                    Open "buddyref/" & Mid(msg, 5) & ".txt" For Input As #filenum4
                    Line Input #filenum4, numusers
                    For I = 1 To numusers
                        Line Input #filenum4, temp
                        If temp <> Username(Index) Then
                            buddylist = buddylist & temp & vbCrLf
                        End If
                        DoEvents
                    Next
                    Close (filenum4)
                End If
                filenum4 = FreeFile
                Open "buddyref/" & Mid(msg, 5) & ".txt" For Output As #filenum4
                
                Write #filenum4, numusers - 1
                If numusers - 1 > 0 Then
                    Print #filenum4, Left(buddylist, Len(buddylist) - 1)
                End If
                Close (filenum4)
                'This creates a new user
            ElseIf Left(msg, 4) = "cre-" Then
                pos = InStr(5, msg, "-")
                suser = Mid(msg, 5, (pos) - 5)
                pass = Mid(msg, pos + 1)
                
                Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = '" & suser & "'")
                If rs.RecordCount = 0 Then
                    rs.AddNew
                    rs!Username = suser
                    rs!Password = pass
                    rs.Update
                    win(Index).SendData "msg-Account created!"
                Else
                    win(Index).SendData "msg-Username already taken!"
                End If
                'Someone requested info on someone else...sends it
            ElseIf Left(msg, 5) = "ginf-" Then
                pos = InStr(1, msg, "-")
                suser = Mid(msg, pos + 1)
                
                If Dir("buddyinfo\" & suser & ".txt") <> "" Then
                    filenum5 = FreeFile
                    Open "buddyinfo\" & suser & ".txt" For Input As #filenum5
                    
                    Do While Not EOF(filenum5)
                        Line Input #filenum5, temp
                        info = info & temp & vbCrLf
                    Loop
                    win(Index).SendData "ginf-" & suser & "-" & info
                    Close (filenum5)
                Else
                    win(Index).SendData "ginf-" & suser & "-" & "User doesn't exist or has no profile!"
                End If
            ElseIf Left(msg, 5) = "sinf-" Then
                info = Mid(msg, 6)
                If info <> "" Then
                    filenum6 = FreeFile
                    Open "buddyinfo/" & Username(Index) & ".txt" For Output As #filenum6
                    Print #filenum6, info
                    Close (filenum6)
                Else
                    If Dir("buddyinfo/" & Username(Index) & ".txt") <> "" Then
                        Kill "buddyinfo/" & Username(Index) & ".txt"
                    End If
                End If
            ElseIf Left(msg, 6) = "forsn-" Then
                realsn = Mid(msg, 7)
                Set rs = db.OpenRecordset("SELECT * FROM Users WHERE Username = '" & realsn & "'")
                win(Index).SendData "forsn-" & rs.Fields(0)
        End If
    Next
End Sub

  
  ' SERVER PER LA CHAT '
  
   
