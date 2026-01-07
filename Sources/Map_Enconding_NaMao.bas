Function ConverteNaMao(ByVal texto As String)

Dim aCaracBugados
Dim aCaracCertos
Dim x As Integer


aCaracBugados = Array(135, 228, 161, 132, 130, 160, 148, 162, 136, 147, 163, 147, 198, 131, 167, 224, 210, 214, 181, 199, 128, 142, 128, 229)
aCaracCertos = Array(231, 245, 237, 227, 233, 225, 245, 243, 234, 244, 250, 244, 227, 226, 186, 211, 202, 205, 193, 195, 199, 195, 199, 213)


For x = 0 To 23

    If InStr(1, texto, Chr(aCaracBugados(x)), vbTextCompare) > 0 Then
       texto = Replace(texto, Chr(aCaracBugados(x)), Chr(aCaracCertos(x)))
    End If

Next x



ConverteNaMao = texto



End Function


