Attribute VB_Name = "criptazione"

Public Function Encrypt(ToEncrypt As Variant) As String

    tmpEncrypt = ""
    Encrypt = ""


    For T = 1 To Len(ToEncrypt)
        tmpEncrypt = tmpEncrypt & Chr(Asc(Mid(ToEncrypt, T, 1)) + 128)


        If Len(tmpEncrypt) = 1000 Then

            DoEvents
                Encrypt = Encrypt & tmpEncrypt
                tmpEncrypt = ""
            End If

        Next T

        Encrypt = Encrypt & tmpEncrypt
    End Function


Public Function Decrypt(ToDecrypt As Variant) As String

    tmpDecrypt = ""
    Decrypt = ""


    For T = 1 To Len(ToDecrypt)
        tmpDecrypt = tmpDecrypt & Chr(Asc(Mid(ToDecrypt, T, 1)) - 128)


        If Len(tmpDecrypt) = 1000 Then

            DoEvents
                Decrypt = Decrypt & tmpDecrypt
                tmpDecrypt = ""
            End If

        Next T

        Decrypt = Decrypt & tmpDecrypt
    End Function
