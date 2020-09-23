Attribute VB_Name = "modSmyle"
Const WM_PASTE = &H302
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Sub ChatMessage(ByVal sUser As String, ByVal Frase As String, ByVal sMessage As String)
    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, iCC As Integer
    Dim CharCombo() As String
    Dim ClipboardContents As Variant
    Dim bClipHasImage As Boolean
    
    bClipHasImage = Clipboard.GetFormat(vbCFBitmap) 'If there's an image in the clipboard
    If bClipHasImage Then chat.picBuffer.Picture = Clipboard.GetData 'Store it to picBuffer
    
    With chat.txtchat
        .Locked = False                 'Must be unlocked for SendMessage() to work
        .SelStart = Len(.Text)          'Move cursor to the end to begin the new message
        
        .SelBold = True                 'Username in bold
        If Trim(sMessage) = "" Then
            .SelText = sUser & Frase
        Else
        
            If Trim(Frase) = "" Then
                .SelText = "<" & sUser & ">" & vbCrLf        'Show the user name
            Else
                .SelText = "<" & sUser & "> " & " <" & Frase & ">" & vbCrLf        'Show the user name
            End If
        End If
        
        lStartMessage = Len(.Text) - 1  'Where the new message begins (search starts here
                                  'for the icons)
        
        .SelBold = False                'Message text is not bold
        .SelColor = vbBlack             'Back to basic black
        .SelText = sMessage & vbCrLf    'Show the message with a linebreak
    End With

    For i = 0 To smyle.imgIcon.Count - 1                  'Loop through each icon
        CharCombo = Split(smyle.imgIcon(i).Tag, " ")      'Get the valid character combinations
                                                    '   which should be delimited by spaces
                                                    '   in the .Tag property
                                                    
        For iCC = 0 To UBound(CharCombo)            'Loop through those character combos
        
                                                    'Find where the char combo starts
            lImagePos = InStr(lStartMessage, chat.txtchat.Text, CharCombo(iCC))


            While lImagePos > 0                     'While the char combo is present
                chat.txtchat.SelStart = lImagePos - 1
                chat.txtchat.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                chat.txtchat.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData smyle.imgIcon(i).Picture        'Set the icon in it
                SendMessage chat.txtchat.hWnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                                                    'Find any more of that same icon
                lImagePos = InStr(lImagePos, chat.txtchat.Text, CharCombo(iCC))
            Wend
        Next iCC
    Next i
    
    chat.txtchat.Locked = True                           'Lock the chat back up
    
    If bClipHasImage Then
        Clipboard.SetData chat.picBuffer.Picture         'Put the old clipboard contents back
    Else
        Clipboard.Clear 'If there were none, then clear it.  There's no use in leaving
    End If              '   an icon sitting in there
    
    chat.txtchat.SelStart = Len(chat.txtchat.Text)            'Move the cursor to the end
End Sub

Public Sub SostituisciSmyle(ByVal sMessage As String)
    Dim lImagePos As Long
    Dim lStartMessage As Long
    Dim i As Integer, iCC As Integer
    Dim CharCombo() As String
    Dim ClipboardContents As Variant
    Dim bClipHasImage As Boolean
    
    bClipHasImage = Clipboard.GetFormat(vbCFBitmap) 'If there's an image in the clipboard
    If bClipHasImage Then chat.picBuffer.Picture = Clipboard.GetData 'Store it to picBuffer
    
    lStartMessage = 1
    
    For i = 0 To smyle.imgIcon.Count - 1                  'Loop through each icon
        CharCombo = Split(smyle.imgIcon(i).Tag, " ")      'Get the valid character combinations
                                                    '   which should be delimited by spaces
                                                    '   in the .Tag property
                                                    
        For iCC = 0 To UBound(CharCombo)            'Loop through those character combos
        
                                                    'Find where the char combo starts
            lImagePos = InStr(lStartMessage, chat.Txtsend.Text, CharCombo(iCC))


            While lImagePos > 0                     'While the char combo is present
                chat.Txtsend.SelStart = lImagePos - 1
                chat.Txtsend.SelLength = Len(CharCombo(iCC)) 'Clear the char combo text
                chat.txtchat.SelText = ""
                
                Clipboard.Clear                             'Clear the clipboard (required)
                Clipboard.SetData smyle.imgIcon(i).Picture        'Set the icon in it
                SendMessage chat.Txtsend.hWnd, WM_PASTE, 0, 0    'Paste it to the rtbChat
                
                                                    'Find any more of that same icon
                lImagePos = InStr(lImagePos, chat.Txtsend.Text, CharCombo(iCC))
            Wend
        Next iCC
    Next i
    
    If bClipHasImage Then
        Clipboard.SetData chat.picBuffer.Picture         'Put the old clipboard contents back
    Else
        Clipboard.Clear 'If there were none, then clear it.  There's no use in leaving
    End If              '   an icon sitting in there
    
    'chat.Txtsend.SelStart = Len(chat.txtsend.Text)            'Move the cursor to the end
End Sub
