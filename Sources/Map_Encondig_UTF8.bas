Public Function fnConverterUTF8(ByVal Texto_para_converter As String)
    Dim l As Long, sUTF8 As String
    Dim iChar As Integer
    Dim iChar2 As Integer
    
    For l = 1 To Len(Texto_para_converter)
        iChar = Asc(Mid(Texto_para_converter, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then
            iChar2 = Asc(Mid(Texto_para_converter, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(Texto_para_converter, l + 1, 1))
            iChar3 = Asc(Mid(Texto_para_converter, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    fnConverterUTF8 = sUTF8
End Function


