
Function DesenhaCaixas(col As Integer, lin As Integer, larg As Integer, altura As Integer, texto As String, tipo As String)

 Sheets("Desenho").Select
 
 texto = Trim(texto)
 
 'Bug fix palavra pequena
 If Len(texto) <= 6 Then
    texto = texto + "__"
 End If
 
 If Trim(texto) = "F_PMnuProSiat()" Or Trim(texto) = "F_PMnuMovSiat()" Or Trim(texto) = "F_PMnuRelSiat()" Then
    texto = "MENU SIAC"
 End If
 
 If Trim(texto) = "FP_ContRot()" Or Trim(texto) = "FP_MenuRotSiret()" Then
    texto = "SIRET"
 End If
  



  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, lin, col, larg, altura).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = texto

        
      If tipo = "prg" Then
      
         With Selection.ShapeRange.Fill
          .Visible = msoTrue
          .ForeColor.ObjectThemeColor = msoThemeColorAccent2
          .ForeColor.TintAndShade = 0
          .ForeColor.Brightness = 0.800000006
          .Transparency = 0
          .Solid
         End With
         
      Else
      
         With Selection.ShapeRange.Fill
          .Visible = msoTrue
          .ForeColor.ObjectThemeColor = msoThemeColorAccent1
          .ForeColor.TintAndShade = 0
          .ForeColor.Brightness = 0.800000006
          .Transparency = 0
          .Solid
         End With
         
      End If
  
    
    testedesenho = 1
    
End Function

