Attribute VB_Name = "salva_text"

Option Explicit
Public Function LoadText(FromFile As String) As String
On Error GoTo Handle
'Checking if the file currently exists
If FileExists(FromFile) = False Then MsgBox "File not found. Check if the file is Currently exists.", vbCritical, "Sorry": Exit Function
Dim sTemp As String
    Open FromFile For Input As #1   'Open the file to read
        sTemp = Input(LOF(1), 1)    'Getting the text
    Close #1                        'Closing the file
    LoadText = sTemp
Exit Function
Handle:
MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function
Public Function SaveText(Text As String, FileName As String) As Boolean
On Error GoTo Handle
Dim sTemp As String
    sTemp = Text
    Open FileName For Append As #1  'Opening the file to SaveText
        Print #1, sTemp             'Printing  the text to the file
    Close #1                        'Closing
    If FileExists(FileName) = False Then    'Check whether the file created
        MsgBox "Unexpectd error occured. File could not be saved", vbCritical, "Sorry"
        SaveText = False    'Returns 'False'
    Else
        SaveText = True     'Returns 'True'
    End If
Exit Function
Handle:
    SaveText = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function
Public Function FileExists(FileName As String) As Boolean
'This function checks the existance of a file
On Error GoTo Handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
Handle:
    FileExists = False
End Function
