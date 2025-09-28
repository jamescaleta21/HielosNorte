Attribute VB_Name = "modSecurity"
' clsSecurity - conversión de modSecurity (Semilla, Codificar, DeCodificar)
Option Explicit
Private Const Base64Table As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

' ----------------------
' Calcula la semilla a partir de una clave (string)
' Devuelve "num1,num2"
' ----------------------
Public Function Semilla(ByVal strClave As String) As String
    Dim lngSemilla1 As Long
    Dim lngSemilla2 As Long
    Dim j As Long
    Dim i As Long
    
    lngSemilla1 = 0
    lngSemilla2 = 0
    j = Len(strClave)
    
    For i = 1 To Len(strClave)
        lngSemilla1 = lngSemilla1 + Asc(Mid$(strClave, i, 1)) * i
        lngSemilla2 = lngSemilla2 + Asc(Mid$(strClave, i, 1)) * j
        j = j - 1
    Next i
    
    ' Normalizar (evitar 0)
    If lngSemilla1 = 0 Then lngSemilla1 = 1
    If lngSemilla2 = 0 Then lngSemilla2 = 1
    
    Semilla = LTrim$(Str$(lngSemilla1)) & "," & LTrim$(Str$(lngSemilla2))
End Function

' ----------------------
' Codifica: aplica transformaciones según semilla (string "n1,n2")
' ----------------------
Private Function Codificar(ByVal strCadena As String, ByVal strSemilla As String) As String
    Dim lngIi1 As Long
    Dim lngIi2 As Long
    Dim i As Long, j As Long
    
    lngIi1 = val(Left$(strSemilla, InStr(strSemilla, ",") - 1))
    lngIi2 = val(Mid$(strSemilla, InStr(strSemilla, ",") + 1))
    
    For i = 1 To Len(strCadena)
        lngIi1 = lngIi1 - i
        lngIi2 = lngIi2 + i
        
        If (i Mod 2) = 0 Then
            Mid$(strCadena, i, 1) = Chr$((Asc(Mid$(strCadena, i, 1)) - lngIi1) And &HFF)
        Else
            Mid$(strCadena, i, 1) = Chr$((Asc(Mid$(strCadena, i, 1)) + lngIi2) And &HFF)
        End If
    Next i
    
    Codificar = strCadena
End Function

' ----------------------
' Decodifica: inverso de Codificar
' ----------------------
Private Function DeCodificar(ByVal strCadena As String, ByVal strSemilla As String) As String
    Dim lngIi1 As Long
    Dim lngIi2 As Long
    Dim i As Long
    
    lngIi1 = val(Left$(strSemilla, InStr(strSemilla, ",") - 1))
    lngIi2 = val(Mid$(strSemilla, InStr(strSemilla, ",") + 1))
    
    For i = 1 To Len(strCadena)
        lngIi1 = lngIi1 - i
        lngIi2 = lngIi2 + i
        
        If (i Mod 2) = 0 Then
            Mid$(strCadena, i, 1) = Chr$((Asc(Mid$(strCadena, i, 1)) + lngIi1) And &HFF)
        Else
            Mid$(strCadena, i, 1) = Chr$((Asc(Mid$(strCadena, i, 1)) - lngIi2) And &HFF)
        End If
    Next i
    
    DeCodificar = strCadena
End Function

' ==== PEGAR EN clsSecurity (al final). No borres lo que ya tienes. ====
' Armadura ASCII segura para TXT: Base64
' Agrega dos métodos públicos: CodificarB64 / DeCodificarB64



' --- Público: codifica y devuelve seguro para TXT (Base64) ---
Public Function CodificarB64(ByVal strCadena As String, ByVal strSemilla As String) As String
    Dim bin As String
    bin = Codificar(strCadena, strSemilla)   ' usa tu método actual (puede producir bytes 0–255)
    CodificarB64 = Base64EncodeStr(bin)      ' lo volvemos ASCII seguro
End Function

' --- Público: toma Base64 desde TXT y devuelve texto original ---
Public Function DeCodificarB64(ByVal strBase64 As String, ByVal strSemilla As String) As String
    Dim bin As String
    bin = Base64DecodeStr(strBase64)         ' recupera bytes originales
    DeCodificarB64 = DeCodificar(bin, strSemilla) ' aplica tu inverso
End Function

' -----------------------
' Helpers Base64 (locales a clsSecurity)
' -----------------------

Private Function Base64EncodeStr(ByVal sText As String) As String
    Dim i As Long, n As Long
    Dim outStr As String
    Dim b1 As Long, b2 As Long, b3 As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long

    outStr = ""
    If Len(sText) = 0 Then
        Base64EncodeStr = ""
        Exit Function
    End If

    For i = 1 To Len(sText) Step 3
        b1 = Asc(Mid$(sText, i, 1))
        If i + 1 <= Len(sText) Then b2 = Asc(Mid$(sText, i + 1, 1)) Else b2 = 0
        If i + 2 <= Len(sText) Then b3 = Asc(Mid$(sText, i + 2, 1)) Else b3 = 0

        c1 = b1 \ 4
        c2 = ((b1 And 3) * 16) Or (b2 \ 16)
        c3 = ((b2 And 15) * 4) Or (b3 \ 64)
        c4 = b3 And 63

        outStr = outStr & Mid$(Base64Table, c1 + 1, 1)
        outStr = outStr & Mid$(Base64Table, c2 + 1, 1)

        If i + 1 <= Len(sText) Then
            outStr = outStr & Mid$(Base64Table, c3 + 1, 1)
        Else
            outStr = outStr & "="
        End If

        If i + 2 <= Len(sText) Then
            outStr = outStr & Mid$(Base64Table, c4 + 1, 1)
        Else
            outStr = outStr & "="
        End If
    Next i

    Base64EncodeStr = outStr
End Function

Private Function B64CharVal(ByVal ch As String) As Long
    Dim p As Long
    If ch = "=" Or Len(ch) = 0 Then
        B64CharVal = -1
        Exit Function
    End If
    p = InStr(1, Base64Table, ch, vbBinaryCompare)
    If p = 0 Then
        B64CharVal = -1
    Else
        B64CharVal = p - 1
    End If
End Function

Private Function Base64DecodeStr(ByVal sText As String) As String
    Dim clean As String
    Dim L As Long, i As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long
    Dim out() As Byte, idx As Long

    ' Limpieza
    clean = Replace(sText, vbCr, "")
    clean = Replace(clean, vbLf, "")
    clean = Replace(clean, " ", "")

    L = Len(clean)
    If L = 0 Then
        Base64DecodeStr = ""
        Exit Function
    End If

    ReDim out(((L + 3) \ 4) * 3 - 1)
    idx = 0

    For i = 1 To L Step 4
        c1 = B64CharVal(Mid$(clean, i, 1))
        c2 = B64CharVal(Mid$(clean, i + 1, 1))
        c3 = B64CharVal(Mid$(clean, i + 2, 1))
        c4 = B64CharVal(Mid$(clean, i + 3, 1))

        If c1 < 0 Or c2 < 0 Then Exit For

        out(idx) = ((c1 * 4) Or (c2 \ 16)) And &HFF
        idx = idx + 1

        If c3 >= 0 Then
            out(idx) = (((c2 And &HF) * 16) Or (c3 \ 4)) And &HFF
            idx = idx + 1

            If c4 >= 0 Then
                out(idx) = (((c3 And 3) * 64) Or c4) And &HFF
                idx = idx + 1
            End If
        End If
    Next i

    If idx = 0 Then
        Base64DecodeStr = ""
    Else
        ReDim Preserve out(idx - 1)
        Base64DecodeStr = StrConv(out, vbUnicode) ' bytes -> String VB6 (1 byte = 1 char)
    End If
End Function
' ==== FIN DEL BLOQUE A PEGAR ====

