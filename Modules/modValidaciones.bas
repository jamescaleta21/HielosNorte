Attribute VB_Name = "modValidaciones"
Public Function SoloNumerosPunto(txt As TextBox, KeyAscii As Integer) As Integer
    Select Case KeyAscii
        Case 8  ' Backspace (borrar)
            SoloNumerosPunto = KeyAscii
        Case 13 ' Enter
            SoloNumerosPunto = KeyAscii
        Case 46 ' Punto "."
            If InStr(1, txt.Text, ".") > 0 Then
                ' Ya existe un punto -> anular
                SoloNumerosPunto = 0
            Else
                If Len(txt.Text) = 0 Then
                    ' Si está vacío y presionan punto -> poner "0."
                    txt.Text = "0."
                    txt.SelStart = Len(txt.Text)  ' mover cursor al final
                    SoloNumerosPunto = 0        ' anular el punto que se iba a escribir
                Else
                    SoloNumerosPunto = KeyAscii
                End If
            End If
        
        Case 48 To 57 ' Números (0–9)
            SoloNumerosPunto = KeyAscii
        
        Case Else
            ' Bloquear cualquier otro caracter
            SoloNumerosPunto = 0
    End Select
End Function

' ===== Nueva función: Validar contenido del TextBox después de un pegado =====
Public Sub ValidarSoloNumerosPunto(txt As TextBox)
    Dim I As Integer
    Dim resultado As String
    Dim tienePunto As Boolean
    
    resultado = ""
    tienePunto = False
    
    For I = 1 To Len(txt.Text)
        Select Case Mid$(txt.Text, I, 1)
            Case "0" To "9"
                resultado = resultado & Mid$(txt.Text, I, 1)
            Case "."
                If Not tienePunto Then
                    ' Solo permite un punto
                    If Len(resultado) = 0 Then
                        resultado = "0."
                    Else
                        resultado = resultado & "."
                    End If
                    tienePunto = True
                End If
        End Select
    Next I
    
    ' Reemplazar texto inválido por el validado
    If txt.Text <> resultado Then
        txt.Text = resultado
        txt.SelStart = Len(txt.Text) ' Colocar cursor al final
    End If
End Sub

Public Function SoloNumeros(KeyAscii As Integer) As Integer
    Select Case KeyAscii
        Case 8   ' Backspace (borrar)
            SoloNumeros = KeyAscii
        Case 13  ' Enter
            SoloNumeros = KeyAscii
        Case 48 To 57 ' Números 0–9
            SoloNumeros = KeyAscii
        Case Else
            ' Anula cualquier otro carácter
            SoloNumeros = 0
    End Select
End Function

Public Sub ValidarSoloNumeros(txt As TextBox)
    Dim I As Integer
    Dim resultado As String
    
    resultado = ""
    
    For I = 1 To Len(txt.Text)
        Select Case Mid$(txt.Text, I, 1)
            Case "0" To "9"
                resultado = resultado & Mid$(txt.Text, I, 1)
            ' Cualquier otro carácter lo ignora
        End Select
    Next I
    
    If txt.Text <> resultado Then
        txt.Text = resultado
        txt.SelStart = Len(txt.Text)
    End If
End Sub

Function FormatoFecha(fecha As Date) As String
Dim strFecha As String

strFecha = Trim(Str(Year(fecha)))
strFecha = Trim(strFecha) & Right("00" & Trim(Str(Month(fecha))), 2)
strFecha = Trim(strFecha) & Right("00" & Trim(Str(Day(fecha))), 2)

FormatoFecha = strFecha

End Function
