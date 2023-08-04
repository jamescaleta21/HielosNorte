Attribute VB_Name = "modValidaciones"
Function NumerosyPunto(Letra As Integer) As Boolean
    'función para el ingreso de Números
    
    If (Letra >= 46 And Letra <= 57) Or (Letra = 8) Or (Letra = 13) Then
        NumerosyPunto = False
    Else
        NumerosyPunto = True
    End If
End Function

Function SoloNumeros(Letra As Integer) As Boolean
    'función para el ingreso de Números
    
    If (Letra >= 47 And Letra <= 57) Or (Letra = 13) Or (Letra = 8) Then
        SoloNumeros = False
    Else
        SoloNumeros = True
    End If
End Function


Function FormatoFecha(fecha As Date) As String
Dim strFecha As String

strFecha = Trim(Str(Year(fecha)))
strFecha = Trim(strFecha) & Right("00" & Trim(Str(Month(fecha))), 2)
strFecha = Trim(strFecha) & Right("00" & Trim(Str(Day(fecha))), 2)

FormatoFecha = strFecha

End Function
